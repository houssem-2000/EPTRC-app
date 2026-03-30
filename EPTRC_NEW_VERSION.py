import sys
import os
from datetime import datetime
import re
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QMessageBox, QWidget,
    QVBoxLayout, QHBoxLayout, QGroupBox, QScrollArea,
    QAction, QFileDialog, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QLineEdit, QComboBox,
    QDialog, QDialogButtonBox, QInputDialog, QCheckBox, QFormLayout
)
from PyQt5.QtGui import QIcon, QFont, QIntValidator, QDoubleValidator
from PyQt5.QtCore import Qt

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ---------------- Paths ---------------- #
SRC_XLSX = resource_path("PATRIMOINE_output.xlsx")
OUT_XLSX = resource_path("PATRIMOINE_output.xlsx")
LOGO_PATH = resource_path("EPTRC.jpg")


class InventoryApp(QMainWindow):
    def __init__(self):
        super().__init__()
        # dynamic data loaded from Excel
        self.products = []
        self.headers = []
        self.dark_mode = False

        # internal flag to suppress itemChanged recursion
        self._suppress_item_changed = False
        self.sort_orders = {}
        # load excel (sets self.products and self.headers)
        self.products, self.headers = self.load_from_excel(SRC_XLSX)

        # ensure QUANTITY STOCK exists in headers & products
        if "QUANTITY STOCK" not in self.headers:
            # append stock column if missing
            self.headers.append("QUANTITY STOCK")
            for p in self.products:
                p.setdefault("QUANTITY STOCK", 1)
        self.initUI()

    # --------------- Excel read/write --------------- #
    def load_from_excel(self, file):
        """Return (products_list, headers_list). If file unreadable, return empty dataset with fallback headers."""
        if not os.path.exists(file):
            QMessageBox.warning(self, "Fichier introuvable", f"Le fichier Excel n'a pas été trouvé:\n{file}\n\nL'application démarre avec un tableau vide.")
            fallback = ["N PARC", "DESIGNATION", "VALEUR ACQ", "QUANTITY STOCK"]
            return [], fallback

        try:
            df = pd.read_excel(file, engine="openpyxl")
        except Exception as e:
            QMessageBox.critical(self, "Erreur lecture Excel", f"Impossible de lire {file}:\n{e}")
            return [], ["N PARC", "DESIGNATION", "VALEUR ACQ", "QUANTITY STOCK"]

        headers = list(df.columns)

        if "QUANTITY STOCK" not in headers:
            df["QUANTITY STOCK"] = 1
            headers = list(df.columns)

        products = []
        for _, row in df.iterrows():
            prod = {col: (row[col] if not pd.isna(row[col]) else "") for col in headers}
            try:
                prod["QUANTITY STOCK"] = int(prod.get("QUANTITY STOCK", 1)) if prod.get("QUANTITY STOCK", "") != "" else 1
            except Exception:
                prod["QUANTITY STOCK"] = 1
            if "VALEUR ACQ" in headers:
                try:
                    prod["VALEUR ACQ"] = float(prod.get("VALEUR ACQ", 0) or 0)
                except Exception:
                    prod["VALEUR ACQ"] = 0.0
            products.append(prod)

        return products, headers

    def save_to_excel(self, file):
        """Save current products to Excel preserving header order (self.headers)."""
        if not self.headers:
            QMessageBox.warning(self, "Erreur", "Aucune colonne définie, impossible d'enregistrer.")
            return

        rows = []
        for p in self.products:
            r = {}
            for h in self.headers:
                val = p.get(h, "")
                if h == "QUANTITY STOCK":
                    try:
                        r[h] = int(val)
                    except Exception:
                        r[h] = 0
                elif h == "VALEUR ACQ":
                    try:
                        r[h] = float(val)
                    except Exception:
                        r[h] = 0.0
                else:
                    r[h] = val
            rows.append(r)

        try:
            df = pd.DataFrame(rows, columns=self.headers)
            df.to_excel(file, index=False, engine="openpyxl")
            QMessageBox.information(self, "Succès", f"Fichier sauvegardé :\n{file}")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible d'enregistrer {file}:\n{e}")

    # --------------- UI & Menus --------------- #
    def initUI(self):
        self.setWindowTitle("EPTRC - Inventaire (PATRIMOINE)")
        self.resize(1200, 800)

        if os.path.exists(LOGO_PATH):
            self.setWindowIcon(QIcon(LOGO_PATH))

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(12)

        menubar = self.menuBar()

        # --- Fichier ---
        file_menu = menubar.addMenu("Fichier")
        load_action = QAction("📂 Charger Excel", self)
        save_action = QAction("💾 Sauvegarder", self)
        exit_action = QAction("🚪 Quitter", self)

        load_action.triggered.connect(self.menu_load_excel)
        save_action.triggered.connect(lambda: self.save_to_excel(OUT_XLSX))
        exit_action.triggered.connect(self.close)

        file_menu.addAction(load_action)
        file_menu.addAction(save_action)
        file_menu.addSeparator()
        file_menu.addAction(exit_action)

        # --- Affichage ---
        view_menu = menubar.addMenu("Affichage")
        icon_dark = QIcon.fromTheme("weather-night")
        icon_light = QIcon.fromTheme("weather-clear")
        if icon_dark.isNull() and os.path.exists(LOGO_PATH):
            icon_dark = QIcon(LOGO_PATH)
        if icon_light.isNull() and os.path.exists(LOGO_PATH):
            icon_light = QIcon(LOGO_PATH)
        self.icon_dark = icon_dark
        self.icon_light = icon_light

        self.toggle_theme_action = QAction(self.icon_dark, "🌙 Mode Sombre", self)
        self.toggle_theme_action.setCheckable(True)
        self.toggle_theme_action.setChecked(False)
        self.toggle_theme_action.toggled.connect(self.toggle_theme)
        view_menu.addAction(self.toggle_theme_action)

        # --- Produits ---
        prod_menu = menubar.addMenu("Produits")
        add_prod_action = QAction("✅ Ajouter Produit", self)
        update_prod_action = QAction("✏️ Modifier Produit", self)
        add_col_action = QAction("➕ Ajouter Colonne", self)
        delete_col_action = QAction("🗑️ Supprimer Colonne", self)
        delete_row_action = QAction("🗑️ Supprimer Ligne", self)

        add_prod_action.triggered.connect(self.open_add_dialog)
        update_prod_action.triggered.connect(self.open_update_dialog)
        add_col_action.triggered.connect(self.add_new_column)
        delete_col_action.triggered.connect(self.delete_column)
        delete_row_action.triggered.connect(self.delete_row)

        prod_menu.addAction(add_prod_action)
        prod_menu.addAction(update_prod_action)
        prod_menu.addAction(add_col_action)
        prod_menu.addAction(delete_col_action)
        prod_menu.addAction(delete_row_action)

        # Title
        title = QLabel("🗂️ Inventaire PATRIMOINE")
        title.setFont(QFont("Segoe UI", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)

        # Table with product data: headers from Excel + "Sélection" last column
        cols = list(self.headers) + ["Sélection"]
        self.table = QTableWidget(0, len(cols))
        # disable default QTableWidget sorting — we use our own column-filtered sorting
        self.table.setSortingEnabled(False)
        header = self.table.horizontalHeader()
        header.sectionClicked.connect(self.handleColumnSort)

        self.table.setHorizontalHeaderLabels(cols)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        main_layout.addWidget(self.table)

        # Buttons grouped left at the bottom
                # Restock group
        produit_group = QGroupBox("🗂️ produit")
        produit_group.setStyleSheet("""
    QGroupBox {
        font-size: 14px;
        font-weight: bold;
        color: #0078D7;
        border: 2px solid #0078D7;
        border-radius: 6px;
        margin-top: 12px;
        padding: 10px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        top: 2px;
        padding: 0 5px;
    }
""")
        btn_layout = QHBoxLayout()
        btn_layout.setAlignment(Qt.AlignLeft)  # <- left alignment as requested
        self.save_btn = QPushButton("💾 Sauvegarder")
        self.save_btn.setFixedSize(160, 40)
        self.update_btn = QPushButton("✏️ Modifier Produit")
        self.update_btn.setFixedSize(160, 40)
        self.exit_btn = QPushButton("🚪 Quitter")
        self.exit_btn.setFixedSize(160, 40)
        btn_layout.addWidget(self.save_btn)
        btn_layout.addWidget(self.update_btn)
        btn_layout.addWidget(self.exit_btn)
        produit_group.setLayout(btn_layout)
        main_layout.addWidget(produit_group)
        # Restock group

        # Add product & add column group
        add_group = QGroupBox("➕ Gestion des colonnes et produits")
        add_group.setStyleSheet("""
    QGroupBox {
        font-size: 14px;
        font-weight: bold;
        color: #0078D7;
        border: 2px solid #0078D7;
        border-radius: 6px;
        margin-top: 12px;
        padding: 10px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        top: 2px;
        padding: 0 5px;
    }
""")
        add_layout = QHBoxLayout()
        self.add_btn = QPushButton("✅ Ajouter Produit")
        self.add_btn.setFixedSize(160, 40)
        self.add_col_btn = QPushButton("➕ Ajouter Colonne")
        self.add_col_btn.setFixedSize(160, 40)
        self.del_col_btn = QPushButton("🗑️ Supprimer Colonne")
        self.del_col_btn.setFixedSize(160, 40)
        self.del_row_btn = QPushButton("🗑️ Supprimer Ligne")
        self.del_row_btn.setFixedSize(160, 40)

        add_layout.addWidget(self.add_btn)
        add_layout.addWidget(self.add_col_btn)
        add_layout.addWidget(self.del_col_btn)
        add_layout.addWidget(self.del_row_btn)
        add_layout.addStretch()

        add_group.setLayout(add_layout)
        main_layout.addWidget(add_group)

        # connect signals
        self.save_btn.clicked.connect(lambda: self.save_to_excel(OUT_XLSX))
        self.exit_btn.clicked.connect(self.close)
        self.add_btn.clicked.connect(self.open_add_dialog)
        self.add_col_btn.clicked.connect(self.add_new_column)
        self.update_btn.clicked.connect(self.open_update_dialog)
        self.del_col_btn.clicked.connect(self.delete_column)
        self.del_row_btn.clicked.connect(self.delete_row)

        # connect to itemChanged to handle checkbox single-select
        self.table.itemChanged.connect(self._on_item_changed)

        # populate table & restock combo
        self.load_table()
        self.apply_light_theme()

    # --------------- Themes --------------- #
    def apply_light_theme(self):
        self.dark_mode = False
        self.setStyleSheet("""
            QWidget {
                font-family: 'Segoe UI';
                font-size: 13px;
                background-color: #F4F4F4;
            }
            QLabel {
                color: #333333;
            }
            QLineEdit, QComboBox {
                padding: 5px;
                font-size: 13px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                background-color: white;
            }
            QPushButton {
                background-color: #0078D7;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #005A9E;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #CCCCCC;
                border-radius: 5px;
                margin-top: 12px;
            }
            QTableWidget {
                background: #FFFFFF;
                alternate-background-color: #FAFAFA;
                border: 1px solid #CCCCCC;
                gridline-color: #DDDDDD;
                color: #111111;
            }
            QHeaderView::section {
                background: #0078D7;
                color: white;
                font-weight: bold;
                padding: 6px;
                border: none;
            }
        """)

    def apply_dark_theme(self):
        self.dark_mode = True
        self.setStyleSheet("""
        QWidget {
            font-family: 'Segoe UI';
            font-size: 13px;
            background-color: #121212;
        }
        QLabel {
            color: #EAEAEA;
        }
        QLineEdit, QComboBox {
            padding: 6px;
            font-size: 13px;
            border: 1px solid #444;
            border-radius: 6px;
            background-color: #1E1E1E;
            color: #EAEAEA;
        }
        QPushButton {
            background-color: #1F6FEB;
            color: white;
            border: none;
            border-radius: 8px;
            padding: 8px 16px;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #1158C7;
        }
        QGroupBox {
            font-weight: bold;
            border: 1px solid #333;
            border-radius: 6px;
            margin-top: 12px;
        }
        QTableWidget {
            background: #1E1E1E;
            alternate-background-color: #252526;
            border: 1px solid #333;
            gridline-color: #444;
            color: #EAEAEA;
        }
        QHeaderView::section {
            background: #1F6FEB;
            color: white;
            font-weight: bold;
            padding: 6px;
            border: none;
        }
    """)

    def toggle_theme(self, checked=False):
        if checked:
            self.apply_dark_theme()
            if hasattr(self, "toggle_theme_action"):
                self.toggle_theme_action.setText("🌞 Mode Clair")
                if not getattr(self, "icon_light", QIcon()).isNull():
                    self.toggle_theme_action.setIcon(self.icon_light)
        else:
            self.apply_light_theme()
            if hasattr(self, "toggle_theme_action"):
                self.toggle_theme_action.setText("🌙 Mode Sombre")
                if not getattr(self, "icon_dark", QIcon()).isNull():
                    self.toggle_theme_action.setIcon(self.icon_dark)

    # --------------- Table population & sync --------------- #
    def load_table(self):
        """Populate the QTableWidget from self.products and self.headers."""
        cols = list(self.headers) + ["Sélection"]
        self.table.blockSignals(True)
        try:
            self.table.setColumnCount(len(cols))
            self.table.setHorizontalHeaderLabels(cols)
            self.table.setRowCount(0)

            # restock combo fill (unique N PARC values)
            
            parc_set = []
            for i, prod in enumerate(self.products):
                self.table.insertRow(i)

                # fill columns from headers (start at column 0)
                for c, h in enumerate(self.headers):
                    value = prod.get(h, "")
                    item = QTableWidgetItem(str(value))
                    if h == "QUANTITY STOCK":
                        item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    else:
                        item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    self.table.setItem(i, c, item)

                # Checkbox as a checkable QTableWidgetItem in the last column
                chk_item = QTableWidgetItem()
                chk_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                chk_item.setCheckState(Qt.Unchecked)
                chk_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(i, len(self.headers), chk_item)

                # Add N PARC to restock combo if present and non-empty

            self.table.resizeColumnsToContents()
            self.table.horizontalHeader().setStretchLastSection(True)
        finally:
            self.table.blockSignals(False)
        self.enable_custom_sorting()
    
    def enable_custom_sorting(self):
        """Keep Qt built-in sorting disabled. Header click is connected once in initUI
        and will call handleColumnSort which sorts the data model and reloads the table."""
        # Ensure built-in sorting is off (we use our custom sorter)
        self.table.setSortingEnabled(False)


    def handle_sort_request(self, column):
        header_name = self.table.horizontalHeaderItem(column).text()
        allowed_cols = ["DATE", "DESIGNATION", "FAMILLE", "SOUS FAMILLE"]
        if header_name not in allowed_cols:
            QMessageBox.information(self, "Tri non autorisé", f"Le tri n'est pas activé pour la colonne '{header_name}'.")
            return

        # gather current table-state (text + flags + checkstate if any)
        table_snapshot = []
        for r in range(self.table.rowCount()):
            row_cells = []
            for c in range(self.table.columnCount()):
                itm = self.table.item(r, c)
                if itm is None:
                    row_cells.append(None)
                else:
                    # make a new item copying useful properties
                    new_itm = QTableWidgetItem(itm.text())
                    new_itm.setFlags(itm.flags())
                    # preserve check state if user-checkable
                    if itm.flags() & Qt.ItemIsUserCheckable:
                        new_itm.setCheckState(itm.checkState())
                    row_cells.append(new_itm)
            table_snapshot.append(row_cells)

        # build sort-key list from snapshot for the requested column
        rows_with_key = []
        date_pattern = re.compile(r'^\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s*(AM|PM)?$', re.IGNORECASE)

        def row_key(row_cells):
            itm = row_cells[column]
            txt = itm.text().strip() if itm else ""
            if date_pattern.match(txt):
                try:
                    txt_norm = " ".join(txt.split())
                    return datetime.strptime(txt_norm, "%m/%d/%Y %I:%M:%S %p")
                except Exception:
                    try:
                        return datetime.strptime(txt, "%m/%d/%Y  %I:%M:%S %p")
                    except Exception:
                        return datetime.min
            return txt.lower()

        ascending = not getattr(self, "_last_sort_asc", True)
        self._last_sort_asc = ascending
        sorted_indices = sorted(range(len(table_snapshot)), key=lambda i: row_key(table_snapshot[i]), reverse=not ascending)

        # now write snapshot rows back into table in sorted order
        self.table.setSortingEnabled(False)
        # clear table rows and rebuild preserving row count
        for r, src_idx in enumerate(sorted_indices):
            for c in range(self.table.columnCount()):
                itm = table_snapshot[src_idx][c]
                # always set a fresh item (avoid moving item objects)
                if itm is not None:
                    self.table.setItem(r, c, itm)
                else:
                    self.table.setItem(r, c, QTableWidgetItem(""))
        self.table.setSortingEnabled(False)
            
    def handleColumnSort(self, column_index):
        """
        Sort self.products by the header clicked.
        Toggles ascending/descending each click (first click = ascending).
        Supports numeric columns, date columns, and text columns.
        """
        # header text (guard against None)
        header_item = self.table.horizontalHeaderItem(column_index)
        if header_item is None:
            return
        header_name = header_item.text().strip()

        # Which columns should be treated as numeric and which as dates?
        # Adjust these sets if your Excel uses different exact header names.
        numeric_headers = {h.upper() for h in ("N PARC", "VALEUR ACQ", "QUANTITY STOCK")}
        date_headers = {h.upper() for h in ("DATE ACQ", "DATE")}  # extend if needed

        # determine column type
        h_up = header_name.upper()
        is_numeric = h_up in numeric_headers
        is_date = h_up in date_headers

        # toggle asc/desc: first click => ascending
        prev = self.sort_orders.get(header_name, None)
        if prev is None:
            asc = True
        else:
            asc = not prev
        self.sort_orders[header_name] = asc

        # helper: parse date robustly
        def parse_date(txt):
            if txt is None or str(txt).strip() == "":
                return datetime.min
            s = str(txt).strip()
            s_norm = " ".join(s.split())
            # try several common formats
            fmts = [
                "%m/%d/%Y %I:%M:%S %p",  # "7/15/2017 12:00:00 AM"
                "%m/%d/%Y %H:%M:%S",
                "%d/%m/%Y %H:%M:%S",
                "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d",
                "%m/%d/%Y",
                "%d/%m/%Y",
            ]
            for f in fmts:
                try:
                    return datetime.strptime(s_norm, f)
                except Exception:
                    continue
            # fallback: try to extract digits and build date, else min
            return datetime.min

        # define key function that returns something sortable
        def keyfunc(prod):
            val = prod.get(header_name, "")
            if is_numeric:
                try:
                    # allow commas and spaces in numbers
                    s = str(val).replace(",", "").strip()
                    if s == "":
                        return float("-inf") if asc else float("inf")
                    return float(s)
                except Exception:
                    # fallback to text compare if conversion fails
                    return str(val)
            elif is_date:
                try:
                    return parse_date(val)
                except Exception:
                    return datetime.min
            else:
    # text (string)
              s = str(val).strip().lower()
              if s == "":
             # Empty goes to bottom ALWAYS
            #  -> if ascending: biggest value
            #  -> if descending: biggest too, reversed order keeps them last
                 return chr(0xffff)  
              return s

        # perform sort on the data model
        try:
            self.products.sort(key=keyfunc, reverse=not asc)
        except Exception as e:
            QMessageBox.warning(self, "Erreur de tri", f"Erreur lors du tri: {e}")
            return

        # reload table with sorted data
        self.load_table()


    def _on_item_changed(self, item: QTableWidgetItem):
        """Ensure only one checkbox is checked at a time. This is called for any item change;
        we only act when the changed item is in the 'Sélection' column and has check state."""
        if self._suppress_item_changed:
            return

        if item is None:
            return

        col = item.column()
        if col != self.table.columnCount() - 1:
            return

        if not (item.flags() & Qt.ItemIsUserCheckable):
            return

        if item.checkState() == Qt.Checked:
            self._suppress_item_changed = True
            try:
                for r in range(self.table.rowCount()):
                    if self.table.item(r, col) is None:
                        continue
                    if r == item.row():
                        continue
                    other = self.table.item(r, col)
                    if other.checkState() == Qt.Checked:
                        other.setCheckState(Qt.Unchecked)
                self.table.selectRow(item.row())
            finally:
                self._suppress_item_changed = False
        else:
            self.table.clearSelection()

    def sync_table_to_products(self):
        """Read stock values from the table and sync to self.products."""
        try:
            stock_idx = self.headers.index("QUANTITY STOCK")
        except ValueError:
            stock_idx = None

        for i, prod in enumerate(self.products):
            if stock_idx is not None:
                stock_item = self.table.item(i, stock_idx)
                if stock_item:
                    try:
                        stock_val = int(stock_item.text())
                        prod["QUANTITY STOCK"] = max(stock_val, 0)
                    except Exception:
                        prod["QUANTITY STOCK"] = 1

    # --------------- Restock, Add/Update product, add column --------------- #
   

    def add_new_column(self):
        col_name, ok = QInputDialog.getText(self, "Nouvelle Colonne", "Nom de la colonne :")
        if not ok or not col_name.strip():
            return
        col_name = col_name.strip()

        if col_name.lower() in [h.lower() for h in self.headers]:
            QMessageBox.warning(self, "Erreur", f"La colonne '{col_name}' existe déjà.")
            return

        self.headers.append(col_name)
        for p in self.products:
            p[col_name] = ""

        self.load_table()
        QMessageBox.information(self, "Succès", f"Colonne '{col_name}' ajoutée avec succès ✅")

    def delete_column(self):
        if not self.headers:
            QMessageBox.warning(self, "Erreur", "Aucune colonne à supprimer.")
            return

        col_name, ok = QInputDialog.getItem(
            self, "Supprimer Colonne",
            "Sélectionnez la colonne à supprimer :",
            self.headers, 0, False
        )
        if not ok or not col_name:
            return

        if col_name in ["N PARC", "DESIGNATION", "VALEUR ACQ", "QUANTITY STOCK"]:
            QMessageBox.warning(self, "Protection", f"La colonne '{col_name}' est protégée et ne peut pas être supprimée.")
            return

        confirm = QMessageBox.question(
            self, "Confirmation",
            f"Voulez-vous vraiment supprimer la colonne '{col_name}' ?",
            QMessageBox.Yes | QMessageBox.No
        )
        if confirm != QMessageBox.Yes:
            return

        if col_name in self.headers:
            self.headers.remove(col_name)
            for p in self.products:
                p.pop(col_name, None)

            self.load_table()
            QMessageBox.information(self, "Succès", f"Colonne '{col_name}' supprimée avec succès ✅")
        else:
            QMessageBox.warning(self, "Erreur", f"Colonne '{col_name}' introuvable.")

    def delete_row(self):
        selected_rows = set(idx.row() for idx in self.table.selectedIndexes())
        if not selected_rows:
            QMessageBox.warning(self, "Erreur", "Veuillez sélectionner au moins une ligne à supprimer.")
            return

        confirm = QMessageBox.question(
            self, "Confirmation",
            f"Voulez-vous vraiment supprimer {len(selected_rows)} ligne(s) ?",
            QMessageBox.Yes | QMessageBox.No
        )
        if confirm != QMessageBox.Yes:
            return

        for row in sorted(selected_rows, reverse=True):
            if row < len(self.products):
                del self.products[row]
                self.table.removeRow(row)

        QMessageBox.information(self, "Succès", "Ligne(s) supprimée(s) avec succès ✅")

    def open_add_dialog(self):
        dlg = ProductDialog(parent=self, title="Ajouter un nouveau produit", headers=self.headers)
        if dlg.exec_() == QDialog.Accepted:
            new_prod = dlg.get_product()

            if "N PARC" in self.headers:
                for p in self.products:
                    if str(p.get("N PARC", "")).strip() == str(new_prod.get("N PARC", "")).strip() and new_prod.get("N PARC", "") != "":
                        QMessageBox.warning(self, "Erreur", f"Produit avec N° Parc '{new_prod.get('N PARC')}' existe déjà.")
                        return

            for h in self.headers:
                new_prod.setdefault(h, "")

            self.products.append(new_prod)
            self.load_table()
            QMessageBox.information(self, "Succès", "Produit ajouté avec succès ✅")

    def open_update_dialog(self):
        selection_col = self.table.columnCount() - 1
        checked_rows = []
        for r in range(self.table.rowCount()):
            itm = self.table.item(r, selection_col)
            if isinstance(itm, QTableWidgetItem) and itm.checkState() == Qt.Checked:
                checked_rows.append(r)

        if len(checked_rows) != 1:
            QMessageBox.warning(self, "Erreur", "Veuillez sélectionner une seule ligne à modifier (une seule case cochée).")
            return

        row = checked_rows[0]
        prod = self.products[row]
        dlg = ProductDialog(parent=self, title="Modifier produit", product=prod.copy(), headers=self.headers)
        if dlg.exec_() == QDialog.Accepted:
            updated_prod = dlg.get_product()
            for h in self.headers:
                updated_prod.setdefault(h, "")

            self.products[row] = updated_prod

            for c, h in enumerate(self.headers):
                item = self.table.item(row, c)
                if item:
                    item.setText(str(updated_prod.get(h, "")))
                else:
                    self.table.setItem(row, c, QTableWidgetItem(str(updated_prod.get(h, ""))))

            QMessageBox.information(self, "Succès", "Produit mis à jour avec succès ✅")
            sel_item = self.table.item(row, selection_col)
            if isinstance(sel_item, QTableWidgetItem):
                sel_item.setCheckState(Qt.Unchecked)

    # --------------- Menu callbacks --------------- #
    def menu_load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choisir fichier Excel", os.path.dirname(SRC_XLSX), "Excel Files (*.xlsx *.xls)")
        if not path:
            return
        prods, hdrs = self.load_from_excel(path)
        if hdrs:
            self.products = prods
            self.headers = hdrs
            if "QUANTITY STOCK" not in self.headers:
                self.headers.append("QUANTITY STOCK")
                for p in self.products:
                    p.setdefault("QUANTITY STOCK", 1)
            self.load_table()

            QMessageBox.information(self, "Chargement terminé", f"Chargement terminé depuis\n{path}")

    # --------------- End class --------------- #


class ProductDialog(QDialog):
    def __init__(self, parent=None, title="Produit", product=None, headers=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumSize(500, 600)
        self.product = product or {}
        self.headers = list(headers) if headers else []
        if "QUANTITY STOCK" not in self.headers:
            self.headers.append("QUANTITY STOCK")
        self.inputs = {}
        self.initUI()

    def initUI(self):
        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        scroll_widget = QWidget()
        form_layout = QFormLayout(scroll_widget)

        for f in self.headers:
            label = f
            line_edit = QLineEdit()
            if f == "VALEUR ACQ":
                line_edit.setValidator(QDoubleValidator(0, 1e12, 2))
            elif f == "QUANTITY STOCK":
                line_edit.setValidator(QIntValidator(0, 999999))

            if self.product:
                val = self.product.get(f, "")
                if val is None:
                    val = ""
                line_edit.setText(str(val))

            form_layout.addRow(label + ":", line_edit)
            self.inputs[f] = line_edit

        scroll_area.setWidget(scroll_widget)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(scroll_area)
        layout.addWidget(btn_box)

    def get_product(self):
        prod = {}
        for f, widget in self.inputs.items():
            text = widget.text().strip()
            if f == "QUANTITY STOCK":
                try:
                    prod[f] = max(int(text), 0)
                except Exception:
                    prod[f] = 1
            elif f == "VALEUR ACQ":
                try:
                    prod[f] = float(text)
                except Exception:
                    prod[f] = 0.0
            else:
                prod[f] = text
        return prod


# ---------------- Main ---------------- #
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = InventoryApp()
    window.show()
    sys.exit(app.exec_())
     