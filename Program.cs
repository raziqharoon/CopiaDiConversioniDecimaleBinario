using System;
using System.Drawing;
using ClosedXML.Excel; // Richiede il pacchetto NuGet ClosedXML

class Program
{
    static void Main()
    {
        string filePath = "C:/Documenti/ConvertitoreDecBin.xlsx";

        // 1. RICHIESTA DEL NUMERO DECIMALE
        int numeroDecimale = 0;
        Console.WriteLine("=== CONVERTITORE DA DECIMALE A BINARIO ===");
        do
        {
            Console.Write("Inserisci un numero decimale tra 0 e 255: ");
            if (!int.TryParse(Console.ReadLine(), out numeroDecimale) || numeroDecimale < 0 || numeroDecimale > 255)
            {
                Console.WriteLine("Valore non valido. Inserisci un numero compreso tra 0 e 255.");
            }
        } while (numeroDecimale < 0 || numeroDecimale > 255);

        // 2. CREAZIONE E SCRITTURA DEL FILE EXCEL
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("DecimaleBinario");

            // Salviamo il numero decimale inserito nella cella A3
            worksheet.Cell(3, 1).Value = numeroDecimale;

            // Formattazione intestazione (Riga 1)
            var headerRow = worksheet.Row(1);
            headerRow.Style.Font.Bold = true; 
            headerRow.Style.Fill.BackgroundColor = XLColor.Yellow; 
            headerRow.Style.Font.FontColor = XLColor.Blue; 

            // Intestazioni delle colonne (Bit da 2^7 a 2^0)
            worksheet.Cell(1, 1).Value = "2^7 (128)";
            worksheet.Cell(1, 2).Value = "2^6 (64)";
            worksheet.Cell(1, 3).Value = "2^5 (32)";
            worksheet.Cell(1, 4).Value = "2^4 (16)";
            worksheet.Cell(1, 5).Value = "2^3 (8)";
            worksheet.Cell(1, 6).Value = "2^2 (4)";
            worksheet.Cell(1, 7).Value = "2^1 (2)";
            worksheet.Cell(1, 8).Value = "2^0 (1)";


            worksheet.Cell(2, 1).FormulaA1 = "=MOD(QUOTIENT($A$3, 128), 2)";
            worksheet.Cell(2, 2).FormulaA1 = "=MOD(QUOTIENT($A$3, 64), 2)";
            worksheet.Cell(2, 3).FormulaA1 = "=MOD(QUOTIENT($A$3, 32), 2)";
            worksheet.Cell(2, 4).FormulaA1 = "=MOD(QUOTIENT($A$3, 16), 2)";
            worksheet.Cell(2, 5).FormulaA1 = "=MOD(QUOTIENT($A$3, 8), 2)";
            worksheet.Cell(2, 6).FormulaA1 = "=MOD(QUOTIENT($A$3, 4), 2)";
            worksheet.Cell(2, 7).FormulaA1 = "=MOD(QUOTIENT($A$3, 2), 2)";
            worksheet.Cell(2, 8).FormulaA1 = "=MOD(QUOTIENT($A$3, 1), 2)";

            // Concatenazione opzionale in un'unica stringa binaria nella cella J1
            worksheet.Cell(1, 10).Value = "Binario Finale";
            worksheet.Cell(2, 10).FormulaA1 = "=A2 & B2 & C2 & D2 & E2 & F2 & G2 & H2";

            worksheet.Columns().AdjustToContents();

            // Salva il file
            workbook.SaveAs(filePath);
            Console.WriteLine($"\nIl File '{filePath}' è stato creato con successo!\n");
        }

        // 3. LETTURA DEL FILE EXCEL CREATO
        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet("DecimaleBinario");

            Console.WriteLine("--- Lettura dal file Excel ---");
            int decLetto = worksheet.Cell(3, 1).GetValue<int>();
            string binarioLetto = worksheet.Cell(2, 10).GetString();

            Console.WriteLine($"Numero Decimale inserito: {decLetto}");
            Console.WriteLine($"Numero Binario calcolato da Excel: {binarioLetto}");
        }
    }
}