import pandas as pd
import re


filename = "excel_file_spare.xlsx"

option_code_inverter = pd.read_excel(filename, sheet_name="Option Code_inverter")
ref = pd.read_excel(filename, sheet_name="Ref")
mvps_spare = pd.read_excel(filename, sheet_name="MVPS spare", skiprows=1, header=0)
inverter_spare = pd.read_excel(filename, sheet_name="Inverter Spare",skiprows=2, header=0)

# print(option_code_inverter.head()) 
# print(ref.head()) 
# print(mvps_spare.head()) 
# print(inverter_spare.head()) 

# Extract option codes (e.g. ['30_6'])
selected_option_codes = option_code_inverter["Final_Result"].dropna().tolist()
print(selected_option_codes)


def evaluate_expression(expr, selected_codes):
    # Repalce option codes with python checks
    def replace_code(match):
        """
        Helper Function
        "30_6 and (3_7 or 2_2)" --> '30_6' in selected_codes and ('3_7' in selected_codes or '2_2' in selected_codes)
        """
        code = match.group(0)
        return f"'{code}' in selected_codes"
    
    # Replace logical operators and format expression
    expr = expr.replace("AND", "and").replace("OR", "or").replace("NOT", "not")
    expr = expr.replace("and"," and ").replace("or", " or ").replace("not", " not ")

    # Replace option codes using regex
    expr = re.sub(r'\b\d+_[\w]+\b', replace_code, expr)

    try:
        return eval(expr)
    except Exception as e:

        print(f"Error evaluating expression: {expr}")
        print(f"Error message: {e}")
        return False


def clean_option_expression(expr):
    """
    Clean an Option expression by preserving only valid logical tokens.
    Removes standalone numbers and unnecessary words, and avoids dangling operators.
    """
    if not expr or str(expr).strip().lower() == "nan":
        return None

    expr = str(expr)

    # Tokenize the expression
    tokens = re.findall(r'\b\d+_\w+\b|and|or|not|\(|\)', expr, flags=re.IGNORECASE)

    # Remove any dangling logical operators (e.g. or or and without operand)
    cleaned_tokens = []
    previous_was_operator = True  # So expression doesn't start with 'and' or 'or'

    for token in tokens:
        token_lower = token.lower()

        if token_lower in {"and", "or"}:
            if previous_was_operator:
                continue  # Skip duplicated or leading operators
            cleaned_tokens.append(token_lower)
            previous_was_operator = True
        else:
            cleaned_tokens.append(token)
            previous_was_operator = False

    # Final cleanup to avoid ending on an operator
    while cleaned_tokens and cleaned_tokens[-1] in {"and", "or", "not"}:
        cleaned_tokens.pop()

    return ' '.join(cleaned_tokens)


# Fallback for lazy header handling
option_column_name = None
for col in mvps_spare.columns:
    if str(col).strip().lower() == "option":
        option_column_name = col
        break

if not option_column_name:
    raise ValueError("Couldnt find the Option column in MVPS spare sheet")

matched_spares = []

for idx, row in mvps_spare.iterrows():
    raw_expr = str(row[option_column_name]).strip()

    # Rule: empty Option = auto-match
    if not raw_expr or raw_expr.lower() == "nan":
        matched_spares.append(row)
        continue
    cleaned_expr = clean_option_expression(raw_expr)


    if not cleaned_expr:
        print(f"Skipping invalid expression at index {idx}: {raw_expr}")
        continue

    # Evaluate cleaned expression
    if evaluate_expression(cleaned_expr, selected_option_codes):
        matched_spares.append(row)


matched_spares_df = pd.DataFrame(matched_spares)

print(mvps_spare.columns)
print(matched_spares_df)

matched_spares_df.to_excel("matched_spares.xlsx", index=False, sheet_name="Matched Spares")