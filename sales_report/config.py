"""Settings for your business: branches, product list and the column names your systems use."""

# Branch name -> words that identify it in a file name.
BRANCHES = {
    "Amsterdam Centrum": ["centrum"],
    "Amsterdam Oost": ["oost"],
    "Utrecht": ["utrecht"],
    "Rotterdam": ["rotterdam"],
    "Den Haag": ["denhaag", "den haag", "den-haag"],
    "Haarlem": ["haarlem"],
}

# The official product names. Typos and different spellings are matched to these.
PRODUCTS = ["Espresso", "Americano", "Cappuccino", "Latte", "Flat White", "Tea", "Hot Chocolate",
            "Croissant", "Banana Bread", "Sandwich"]

# Other names for the same product (e.g. the Dutch names used by one branch).
PRODUCT_ALIASES = {
    "koffie verkeerd": "Latte", "thee": "Tea", "warme chocolademelk": "Hot Chocolate",
    "bananenbrood": "Banana Bread", "broodje": "Sandwich",
}

# Column headers in the exports -> the standard column. Matching ignores case and symbols.
COLUMNS = {
    "date": ["date", "datum", "sale date", "transaction date", "day", "dag"],
    "product": ["product", "item", "product name", "artikel", "description", "omschrijving"],
    "qty": ["qty", "quantity", "units", "aantal", "qty sold", "pcs"],
    "price": ["price", "unit price", "prijs", "price (€)", "price eur"],
    "revenue": ["total", "revenue", "omzet", "amount", "sales", "turnover"],
}

COMPANY = "Harbor & Pine Cafés"
