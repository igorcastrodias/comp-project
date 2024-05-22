import pandas as pd
import numpy as np
import time
from datetime import datetime

# Função para ler o arquivo Excel
def read_excel_file(file_path):
    return pd.read_excel(file_path)

# Função para processar os dados e criar matrizes
def process_data(data_frame, matrix_size=12):
    row_count = data_frame.shape[0] - 1  # Ignorar a linha de cabeçalho
    col_count = data_frame.shape[1] - 1  # Ignorar a primeira coluna

    matrices = []
    for start_row in range(1, row_count - matrix_size + 2):
        matrix_data = np.zeros((matrix_size, matrix_size))
        for i in range(matrix_size):
            for j in range(matrix_size):
                value = data_frame.iloc[start_row + i, j + 1]
                try:
                    matrix_data[i, j] = float(value)
                except ValueError:
                    print(f"Valor inválido na linha {start_row + i + 1}, coluna {j + 2}. Substituído por 0.")
                    matrix_data[i, j] = 0
        matrices.append(matrix_data)
    return matrices

# Função para calcular inversas, exportar resultados e verificar determinante
def calculate_and_export_matrices(matrices, matrix_size=12):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"result_{timestamp}.txt"

    with open(file_name, "w") as writer:
        for k, matrix in enumerate(matrices):
            determinant = np.linalg.det(matrix)
            if determinant != 0:
                writer.write(f"Matriz {k + 1}:\n")
                writer.write(f"{matrix}\n\n")

                inverse = np.linalg.inv(matrix)
                writer.write(f"Inversa da Matriz {k + 1}:\n")
                writer.write(f"{inverse}\n\n")

                writer.write(f"Determinante da Matriz {k + 1}: {determinant}\n\n")
            else:
                print(f"Matriz {k + 1} tem determinante igual a 0.")
                print("Matriz:")
                print(matrix)

    print(f"Número de matrizes geradas: {len(matrices)}")

def main():
    file_path = "FECHAMENTO_MAIS_NEGOCIADAS.xlsx"

    start_time = time.time()

    # Ler os dados do Excel
    data_frame = read_excel_file(file_path)
    read_time = time.time()
    print(f"Tempo de leitura do arquivo: {read_time - start_time:.2f} segundos")

    # Processar os dados e criar matrizes
    matrices = process_data(data_frame)
    process_time = time.time()
    print(f"Tempo de criação das matrizes: {process_time - read_time:.2f} segundos")

    # Calcular inversas e exportar resultados
    calculate_and_export_matrices(matrices)
    calculate_time = time.time()
    print(f"Tempo de cálculo das inversas: {calculate_time - process_time:.2f} segundos")

    # Tempo total de execução
    total_time = time.time()
    print(f"Tempo total de execução: {total_time - start_time:.2f} segundos")

if __name__ == "__main__":
    main()
