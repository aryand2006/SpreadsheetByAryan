def sum_formula(values):
    return sum(values)

def average_formula(values):
    return sum(values) / len(values) if values else 0

def count_formula(values):
    return len(values)

def max_formula(values):
    return max(values) if values else None

def min_formula(values):
    return min(values) if values else None

def evaluate_formula(formula, context):
    # This function will parse and evaluate the formula based on the context provided
    # For simplicity, this is a placeholder for actual formula parsing logic
    if formula.startswith("SUM"):
        return sum_formula(context)
    elif formula.startswith("AVERAGE"):
        return average_formula(context)
    elif formula.startswith("COUNT"):
        return count_formula(context)
    elif formula.startswith("MAX"):
        return max_formula(context)
    elif formula.startswith("MIN"):
        return min_formula(context)
    else:
        raise ValueError(f"Unknown formula: {formula}")