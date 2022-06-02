using FLS;
using FLS.Rules;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace System_Fuzzy_Logic
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Excel
            string filePath = "A:\\Aliim\\Semester 4\\AI\\Minggu 7\\AI Kelompok E.xlsx";
            Application excel = new Application();
            Workbook wb = excel.Workbooks.Open(filePath);
            Worksheet ws = wb.Worksheets[1];

            Range cellGaji;
            Range cellIPK;

            // Fuzzification
            var gaji = new LinguisticVariable("Gaji");
            var kecil = gaji.MembershipFunctions.AddTrapezoid("Kecil", 0, 0, 1, 3);
            var sedang = gaji.MembershipFunctions.AddTrapezoid("Sedang", 1, 3, 4, 6);
            var besar = gaji.MembershipFunctions.AddTrapezoid("Besar", 4, 6, 7, 12);
            var sBesar = gaji.MembershipFunctions.AddTrapezoid("Sangat Besar", 7, 12, 14, 14);

            var ipk = new LinguisticVariable("IPK");
            var buruk = ipk.MembershipFunctions.AddTrapezoid("Buruk", 0, 0, 2, 2.75);
            var cukup = ipk.MembershipFunctions.AddTriangle("Cukup", 2, 2.75, 3.25);
            var bagus = ipk.MembershipFunctions.AddTrapezoid("Bagus", 2.75, 3, 3.25, 4);

            var nk = new LinguisticVariable("Nilai Kelayakan");
            var rendah = nk.MembershipFunctions.AddTrapezoid("Rendah", 0, 0, 50, 80);
            var tinggi = nk.MembershipFunctions.AddTrapezoid("Tinggi", 50, 80, 100, 100);

            //Inference
            IFuzzyEngine fuzzyEngine = new FuzzyEngineFactory().Default();

            var rule1 = Rule.If(ipk.Is(buruk).And(gaji.Is(kecil))).Then(nk.Is(rendah));
            var rule2 = Rule.If(ipk.Is(buruk).And(gaji.Is(sedang))).Then(nk.Is(rendah));
            var rule3 = Rule.If(ipk.Is(buruk).And(gaji.Is(besar))).Then(nk.Is(rendah));
            var rule4 = Rule.If(ipk.Is(buruk).And(gaji.Is(sBesar))).Then(nk.Is(rendah));

            var rule5 = Rule.If(ipk.Is(cukup).And(gaji.Is(kecil))).Then(nk.Is(tinggi));
            var rule6 = Rule.If(ipk.Is(cukup).And(gaji.Is(sedang))).Then(nk.Is(rendah));
            var rule7 = Rule.If(ipk.Is(cukup).And(gaji.Is(besar))).Then(nk.Is(rendah));
            var rule8 = Rule.If(ipk.Is(cukup).And(gaji.Is(sBesar))).Then(nk.Is(rendah));

            var rule9 = Rule.If(ipk.Is(bagus).And(gaji.Is(kecil))).Then(nk.Is(tinggi));
            var rule10 = Rule.If(ipk.Is(bagus).And(gaji.Is(sedang))).Then(nk.Is(tinggi));
            var rule11 = Rule.If(ipk.Is(bagus).And(gaji.Is(besar))).Then(nk.Is(tinggi));
            var rule12 = Rule.If(ipk.Is(bagus).And(gaji.Is(sBesar))).Then(nk.Is(rendah));

            fuzzyEngine.Rules.Add(rule1, rule2, rule3, rule4, rule5, rule6, rule7, 
                rule8, rule9, rule10, rule11, rule12);

            // Defuzzification
            cellIPK = ws.Range["C2"];
            cellGaji = ws.Range["D2"];
            var result = fuzzyEngine.Defuzzify(new { ipk = cellIPK.Value, gaji = cellGaji.Value });
            Console.WriteLine("IPK: "+ cellIPK.Value + " |Gaji: " + cellGaji.Value + " = " +result+"\n");

            cellIPK = ws.Range["C29"];
            cellGaji = ws.Range["D29"];
            result = fuzzyEngine.Defuzzify(new { ipk = cellIPK.Value, gaji = cellGaji.Value });
            Console.WriteLine("IPK: " + cellIPK.Value + " |Gaji: " + cellGaji.Value + " = " + result + "\n");

            cellIPK = ws.Range["C81"];
            cellGaji = ws.Range["D81"];
            result = fuzzyEngine.Defuzzify(new { ipk = cellIPK.Value, gaji = cellGaji.Value });
            Console.WriteLine("IPK: " + cellIPK.Value + " |Gaji: " + cellGaji.Value + " = " + result + "\n");

        }

    }
}
