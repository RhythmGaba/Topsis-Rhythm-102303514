import sys
import pandas as pd
import numpy as np

# Check all arguments validation
def check_arguments():
    if len(sys.argv) != 5:
        print("Usage:")
        print("python TOPSIS_code.py <input.xlsx> <weights> <impacts> <output.xlsx>")
        sys.exit(1)

# Read input file
def read_file(filename):
    try:
        return pd.read_excel(filename)
    except:
        print("Error: Input file not found or invalid")
        sys.exit(1)

# Check column count & validation
def validate_data(df):
    if df.shape[1] < 3:
        print("Error: File must have at least 3 columns")
        sys.exit(1)

    for col in df.columns[1:]:
        if not pd.api.types.is_numeric_dtype(df[col]):
            print(f"Error: Column {col} must be numeric")
            sys.exit(1)

# Weights & Impacts with ","
def weights_impacts(weights_str, impacts_str, n):
    weights = [float(i) for i in weights_str.split(',')]
    impacts = impacts_str.split(',')

    if len(weights) != n or len(impacts) != n:
        print("Error: Weights/Impacts count mismatch")
        sys.exit(1)

    for i in impacts:
        if i not in ['+', '-']:
            print("Error: Impacts must be + or - only")
            sys.exit(1)

    return weights, impacts

# Normalization of the data
def normalize_matrix(matrix):
    norm = matrix.copy()
    for j in range(matrix.shape[1]):
        denom = np.sqrt(np.sum(matrix[:, j] ** 2))
        norm[:, j] = matrix[:, j] / denom
    return norm

# Apply weights
def apply_weights(norm_matrix, weights):
    for j in range(len(weights)):
        norm_matrix[:, j] *= weights[j]
    return norm_matrix

# Calculating ideal best & worst
def ideal_solutions(matrix, impacts):
    best = []
    worst = []

    for j in range(len(impacts)):
        if impacts[j] == '+':
            best.append(matrix[:, j].max())
            worst.append(matrix[:, j].min())
        else:
            best.append(matrix[:, j].min())
            worst.append(matrix[:, j].max())

    return np.array(best), np.array(worst)

# Calculating Euclidean Distance &  TOPSIS score
def calculate_scores(matrix, best, worst):
    scores = []
    for i in range(matrix.shape[0]):
        d_pos = np.sqrt(np.sum((matrix[i] - best) ** 2))
        d_neg = np.sqrt(np.sum((matrix[i] - worst) ** 2))
        scores.append(d_neg / (d_pos + d_neg))
    return scores

#MAIN
def main():
    check_arguments()

    input_data_file = sys.argv[1]
    weights = sys.argv[2]
    impacts = sys.argv[3]
    output_result_file = sys.argv[4]

    df = read_file(input_data_file)
    validate_data(df)

    matrix = df.iloc[:, 1:].values
    weights, impacts = weights_impacts(weights, impacts, matrix.shape[1])

    norm_matrix = normalize_matrix(matrix)
    weighted_matrix = apply_weights(norm_matrix, weights)

    best, worst = ideal_solutions(weighted_matrix, impacts)
    scores = calculate_scores(weighted_matrix, best, worst)

    df["Topsis Score"] = scores
    df["Rank"] = df["Topsis Score"].rank(ascending=False).astype(int)

    df.to_excel(output_result_file, index=False)
    print("Result saved successfully in", output_result_file)


if __name__ == "__main__":
    main()
