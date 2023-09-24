# Categorization rules
categorization_rules = [
    {
        'condition': lambda transaction: 'Grocery' in transaction['Description/Transaction Name'],
        'category': 'Groceries'
    },
    {
        'condition': lambda transaction: transaction['Amount'] < 0,
        'category': 'Expense'
    },
    # Add more rules as needed
]
