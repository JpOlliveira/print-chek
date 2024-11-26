using System.Drawing;
using System.Globalization;

const int CARACTERS = 140;
#region Positions
const int SCALE_FIX = -170;
const int DEHIATION = 3350 - SCALE_FIX;

// Preenchimento valor
const int x_valueInReal = 7100;
const int y_valueInReal = 30;
const int width_valueInReal = 1600;
const int height_valueInReal = 210;

// Preenchimento valor extenso
const int x_extensiveValue = 2253;
const int y_extensiveValue = 550;
const int width_extensiveValue = 6820;
const int height_extensiveValue = 210;

// Preenchimento campo nominal
const int x_nominalValue = 1280;
const int y_nominalValue = 1200;
const int width_nominalValue = 6990;
const int height_nominalValue = 210;

// Preenchimento campo cidade
const int x_city = 4415;
const int y_city = 1520;
const int width_city = 1467;
const int height_city = 210;

// Preenchimento campo dia
const int x_day = 5983;
const int y_day = 1520;
const int width_day = 480;
const int height_day = 210;

// Preenchimento campo mês
const int x_month = 6620;
const int y_month = 1520;
const int width_month = 1550;
const int height_month = 210;

// Preenchimento campo ano
const int x_year = 8525;
const int y_year = 1530;
const int width_year = 490;
const int height_year = 210;
#endregion

Image originalChequesImage = Image.FromFile(".\\Scan2024-11-25_145150.png");
Image logo = Image.FromFile(".\\logo-sem-fundo.png");
Brush brush = new SolidBrush(Color.Black);
Font arial = new("Arial", 140, FontStyle.Regular);
var pt_Br = CultureInfo.GetCultureInfo("pt-BR");

object fileCount = 0;
Console.WriteLine($"Cheques detectados na planilha xlsx: {cheques.Count}");
List<Task> savingTasks = [];
SemaphoreSlim semaphore = new(1, 1);
foreach (var chequeGroup in cheques.Chunk(4))
    savingTasks.Add(Task.Run(delegate
    {
        int localCopy = 0;
        semaphore.Wait();
        fileCount = (int)fileCount + 1;
        localCopy = (int)fileCount;
        semaphore.Release();
        Console.WriteLine($"Preparando escrita do arquivo: {localCopy}");
        Bitmap emptyChequesBitmapSketch = new(originalChequesImage.Width, originalChequesImage.Height);
        Graphics graphics = Graphics.FromImage(emptyChequesBitmapSketch);
        Console.WriteLine($"[{localCopy}] Imagem carregada para a memoria");

        int i = 0;
        foreach (var cheque in chequeGroup)
        {
            var dehiation = i switch
            {
                0 => SCALE_FIX,
                1 => (i * DEHIATION) - 70,
                2 => i * DEHIATION,
                3 => (i * DEHIATION) + 70,
                _ => throw null!
            };

            string valueInRealText = $"# {cheque.Value:#,###.00} #";
            Rectangle valueInReal = new(x_valueInReal, (i == 0 ? 130 : 0) + dehiation + y_valueInReal, width_valueInReal, height_valueInReal);
            graphics.DrawString(valueInRealText, arial, brush, valueInReal);

            string extensiveValueText = cheque.Extensive.PadRight(cheque.Extensive.Length + 1) + new string('X', CARACTERS - cheque.Extensive.Length - 1);
            Rectangle extensiveValue = new(x_extensiveValue, dehiation + y_extensiveValue, width_extensiveValue, height_extensiveValue);
            graphics.DrawString(extensiveValueText, arial, brush, extensiveValue);

            string nominalValueText = cheque.Nomenclature;
            Rectangle nominalValue = new(x_nominalValue, dehiation + y_nominalValue, width_nominalValue, height_nominalValue);
            graphics.DrawString(nominalValueText, arial, brush, nominalValue);

            string cityText = cheque.City;
            Rectangle cityValue = new(x_city, dehiation + y_city, width_city, height_city);
            graphics.DrawString(cityText, arial, brush, cityValue);

            string dayText = cheque.Date.Day.ToString();
            Rectangle dayValue = new(x_day, dehiation + y_day, width_day, height_day);
            graphics.DrawString(dayText, arial, brush, dayValue);

            string monthText = cheque.Date.ToString("MMMM", pt_Br);
            Rectangle monthValue = new(x_month, dehiation + y_month, width_month, height_month);
            graphics.DrawString(monthText, arial, brush, monthValue);

            string yearText = cheque.Date.ToString("yy", pt_Br); ;
            Rectangle yearValue = new(x_year, dehiation + y_year, width_year, height_year);
            graphics.DrawString(yearText, arial, brush, yearValue);

            int logoX = 3150;
            int logoY = dehiation + 1780;
            int logoWidth = 1000;
            int logoHeight = 1000;
            semaphore.Wait();
            graphics.DrawImage(logo, new Rectangle(logoX, logoY, logoWidth, logoHeight));
            semaphore.Release();

            i++;
        }
        Directory.CreateDirectory(".\\outs\\");
        Console.WriteLine($"[{localCopy}] processo de salvamento da imagem");
        emptyChequesBitmapSketch.Save($".\\outs\\cheques[{localCopy}].png");
    }))
;
Task.WaitAll([.. savingTasks]);