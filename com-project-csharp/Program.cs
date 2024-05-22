using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using ExcelDataReader;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            // Caminho para o arquivo Excel
            string filePath = "FECHAMENTO_MAIS_NEGOCIADAS.xlsx";
            List<Matrix<double>> matrices = new List<Matrix<double>>();

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            // Ler os dados do Excel
            DataTable dataTable = ReadExcelFile(filePath);
            stopwatch.Stop();
            TimeSpan readTime = stopwatch.Elapsed;

            // Processar os dados e criar matrizes
            stopwatch.Restart();
            matrices = ProcessData(dataTable, out int matrixSize);
            stopwatch.Stop();
            TimeSpan processTime = stopwatch.Elapsed;

            // Calcular inversas e outras operações
            stopwatch.Restart();
            CalculateAndExportMatrices(matrices, matrixSize);
            stopwatch.Stop();
            TimeSpan calculateTime = stopwatch.Elapsed;

            // Tempo total de execução
            TimeSpan totalTime = readTime + processTime + calculateTime;

            // Exibir os tempos
            Console.WriteLine($"Tempo de leitura do arquivo: {readTime}");
            Console.WriteLine($"Tempo de criação das matrizes: {processTime}");
            Console.WriteLine($"Tempo de cálculo das inversas: {calculateTime}");
            Console.WriteLine($"Tempo total de execução: {totalTime}");
        }

        static DataTable ReadExcelFile(string filePath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    return result.Tables[0];
                }
            }
        }

        static List<Matrix<double>> ProcessData(DataTable dataTable, out int matrixSize)
        {
            matrixSize = 12; // Default matrix size
            int rowCount = dataTable.Rows.Count - 1; // Ignorando a linha de cabeçalho
            int colCount = dataTable.Columns.Count - 1; // Ignorando a primeira coluna

            List<Matrix<double>> matrices = new List<Matrix<double>>();

            for (int startRow = 1; startRow <= rowCount - matrixSize + 1; startRow++)
            {
                var matrixData = new double[matrixSize, matrixSize];
                for (int i = 0; i < matrixSize; i++)
                {
                    for (int j = 0; j < matrixSize; j++)
                    {
                        var value = dataTable.Rows[startRow + i][j + 1].ToString();
                        if (double.TryParse(value, out double result))
                        {
                            matrixData[i, j] = result;
                        }
                        else
                        {
                            Console.WriteLine($"Valor inválido na linha {startRow + i + 1}, coluna {j + 2}. Substituído por 0.");
                            matrixData[i, j] = 0;
                        }
                    }
                }
                matrices.Add(DenseMatrix.OfArray(matrixData));
            }

            return matrices;
        }

        static void CalculateAndExportMatrices(List<Matrix<double>> matrices, int matrixSize)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"result_{timestamp}.txt";

            using (StreamWriter writer = new StreamWriter(fileName))
            {
                for (int k = 0; k < matrices.Count; k++)
                {
                    var matrix = matrices[k];
                    if (matrix.Determinant() != 0)
                    {
                        writer.WriteLine($"Matriz {k + 1}:");
                        writer.WriteLine(matrix.ToString(12, 12));
                        writer.WriteLine();

                        var inverse = matrix.Inverse();
                        writer.WriteLine($"Inversa da Matriz {k + 1}:");
                        writer.WriteLine(inverse.ToString(12, 12));
                        writer.WriteLine();

                        writer.WriteLine($"Determinante da Matriz {k + 1}: {matrix.Determinant()}");
                        writer.WriteLine();
                    }
                    else
                    {
                        Console.WriteLine($"Matriz {k + 1} tem determinante igual a 0.");
                        Console.WriteLine("Matriz:");
                        Console.WriteLine(matrix.ToString(12, 12));
                    }
                }
            }

            Console.WriteLine($"Número de matrizes geradas: {matrices.Count}");
        }
    }
}
