
using ClosedXML.Excel;
using System.Diagnostics;

/// <summary>
/// Aplicação de console que adiciona um numero random na primeira linha de arquivo xlsx para que o hash do arquivo seja diferente;;
/// </summary>
class Program
{
    static void Main()
    {
        string sourceDirectory = @"C:\Users\Elian\Documents\Timoneiro\Brusque\2023 completa 2022\DFE_Relatório_CTe_para_MOVEC1";
        string[] files = Directory.GetFiles(sourceDirectory);

        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();
        foreach (string file in files)
        {
            FileInfo fileInfo = new FileInfo(file);

            if (fileInfo.Extension == ".xlsx")
            {
                using (XLWorkbook package = new XLWorkbook(fileInfo.FullName))
                {
                    IXLWorksheet worksheet = package.Worksheet(1); // Primeira planilha

                    Random random = new Random();
                    // Gerar um número aleatório para alterar a primeira célula
                    int randomNumber = random.Next(1000);

                    // Alterar a primeira célula da planilha
                    worksheet.Row(1).Cell(1).Value = randomNumber.ToString() + " " + worksheet.Row(1).Cell(1).Value;

                    // Criar e salvar o novo arquivo triplicado
                    string newFileName = Path.Combine(sourceDirectory, Path.GetFileNameWithoutExtension(file) + "_triplicado" + randomNumber.ToString() + Path.GetExtension(file));
                    package.SaveAs(new FileInfo(newFileName).FullName);

                    // Criar e salvar o segundo arquivo triplicado
                    randomNumber = random.Next(1000);
                    string secondNewFileName = Path.Combine(sourceDirectory, Path.GetFileNameWithoutExtension(file) + "_triplicado2" + randomNumber.ToString() + Path.GetExtension(file));
                    package.SaveAs(new FileInfo(secondNewFileName).FullName);

                    Console.WriteLine($"Arquivo triplicado criado: {newFileName}");
                    Console.WriteLine($"Segundo arquivo triplicado criado: {secondNewFileName}");
                }
            }
            else
            {
                Console.WriteLine($"O arquivo {file} não é um arquivo Excel (.xlsx) e será ignorado.");
            }
        }
        stopwatch.Stop();
        Console.WriteLine($"Processo concluído em { stopwatch.Elapsed.ToString(@"hh\:mm\:ss")}");
    }
}