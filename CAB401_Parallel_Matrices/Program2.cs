using System.Diagnostics;
using System.Reflection;


//logging - Windows only
using Excel = Microsoft.Office.Interop.Excel;


//Before you ask, Arguments.dll is my own implementation,
//and whilst it is included in the project for the purpose of simplifying the argument process
//I do use it in a professional manner outside of Uni, and therefore falls under copyright.
using Arguments;

namespace Cab401_Parallel
{
    public static class ProgramSettings
    {
        //Logging is quite time consuming due to using the excel interop
        [FieldArgument("-log", true)]
        public static bool doLogging = false;

        [MethodArgument("-dimensions", typeof(uint), typeof(uint), typeof(uint), typeof(uint))]
        static void setDims(uint Ah, uint Aw, uint Bh, uint Bw)
        {
            MatrixA_height = Ah;
            MatrixA_width = Aw;
            MatrixB_height = Bh;
            MatrixB_width = Bw;
        }
        public static uint MatrixA_width { get; private set; } = 500;
        public static uint MatrixB_width { get; private set; } = 300;
        public static uint MatrixA_height { get; private set; } = 300;
        public static uint MatrixB_height { get; private set; } = 250;


        [FieldArgument("-iter")]
        public static uint MaxIterations = 100;
    }


    public static class Program2
    {
        public static void Main(string[] args)
        {
            //parsing settings
            ArgumentSystem.ParseArgs(args, typeof(ProgramSettings), BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public);



            Stopwatch stp = new Stopwatch();

            #region fixed tables
            //double[][] MatrixA = new double[][]
            //{
            //    new double[]{  5.78, 31.57, 65.22 },
            //    new double[]{  5.92, 14.72, 68.76 },
            //    new double[]{ 35.94, 49.68, 49.51 },
            //    new double[]{ 52.22, 85.05, 27.97 },
            //    new double[]{ 79.58, 64.84, 90.53 }
            //};
            //double[][] MatrixB = new double[][]
            //{
            //    new double[]{ 85.37, 39.78,  6.52,  0.63 },
            //    new double[]{ 25.20, 38.73, 87.86, 32.89 },
            //    new double[]{ 29.02, 51.95, 10.04, 89.44 }
            //};
            //double[][] Matrix_Result = new double[0][];
            #endregion

            #region Random Setup

            double[,] MatrixA = new double[ProgramSettings.MatrixA_height, ProgramSettings.MatrixA_width];
            double[,] MatrixB = new double[ProgramSettings.MatrixB_height, ProgramSettings.MatrixB_width];
            double[,] Matrix_Result;

            //generate random tables
            for (int i = 0; i < MatrixA.GetLength(0); i++)
                for (int j = 0; j < MatrixA.GetLength(1); j++)
                    MatrixA[i,j] = Rand(0, 100);

            for (int i = 0; i < MatrixB.GetLength(0); i++)
                for (int j = 0; j < MatrixB.GetLength(1); j++)
                    MatrixB[i, j] = Rand(0, 100);
            #endregion

            #region Logging setup
            ExcelHandler? excel = null;
            if (ProgramSettings.doLogging)
            {
                excel = new ExcelHandler();
                excel.setup(false);
            }
            #endregion



            Console.WriteLine("Press any key to begin...");
            //Console.ReadKey(true);


            if (MatrixA.GetLength(0) != MatrixB.GetLength(1))
                throw new Exception("Matrix A and B have incompatible Widths and Heights respectively.");

            Matrix_Result = new double[MatrixA.GetLength(0), MatrixB.GetLength(1)];

            uint Iteration = 0;
            uint MaxIterations = ProgramSettings.MaxIterations;

            //This is cause im reading doLogging each time a process is done
            TimeSpan seq = new TimeSpan(), para = new TimeSpan(), dpara = new TimeSpan();

            while (Iteration < MaxIterations)
            {
                //Sequential
                stp.Start();
                for (int a = 0; a < MatrixA.GetLength(0); a++)
                {
                    for (int b = 0; b < MatrixB.GetLength(1); b++)
                    {
                        double total = 0;
                        for (int c = 0; c < MatrixB.GetLength(0); c++)
                            total += MatrixA[a, c] * MatrixB[c, b];
                        Matrix_Result[a, b] = total;
                    }
                }
                stp.Stop();
                if (ProgramSettings.doLogging)
                    seq = stp.Elapsed;


                stp.Reset();


                //Single Parallel
                stp.Start();
                Parallel.For(0, MatrixA.GetLength(0), a =>
                {
                    for (int b = 0; b < MatrixB.GetLength(1); b++)
                    {
                        double total = 0;
                        for (int c = 0; c < MatrixB.GetLength(0); c++)
                            total += MatrixA[a, c] * MatrixB[c, b];
                        Matrix_Result[a, b] = total;
                    }
                });
                stp.Stop();
                if (ProgramSettings.doLogging)
                    para = stp.Elapsed;


                stp.Reset();


                //Dual Parallel
                stp.Start();
                Parallel.For(0, MatrixA.GetLength(0), a =>
                {
                    Parallel.For(0, MatrixB.GetLength(1), b =>
                    {
                        double total = 0;
                        for (int c = 0; c < MatrixB.GetLength(0); c++)
                            total += MatrixA[a, c] * MatrixB[c, b];
                        Matrix_Result[a, b] = total;
                    });
                });
                stp.Stop();
                if (ProgramSettings.doLogging)
                {
                    dpara = stp.Elapsed;
                    excel.WriteColumn((int)Iteration, seq, para, dpara);
                }

                Iteration++;
            }
            if (ProgramSettings.doLogging)
                excel.SetLock(true);
            Console.WriteLine("Loop complete, press anything to end task...");
            Console.ReadKey();
        }

        public static double Rand(uint min, uint max)
        {
            Random r = new Random();
            return r.NextDouble() * (max - min) + min;
        }
    }

    public class ExcelHandler
    {
        Excel.Application appl;
        Excel.Workbook wrkB;
        Excel.Worksheet wrkS;

        int DataPoint_StartColumn = 2;

        public ExcelHandler()
        {
            appl = new Excel.Application();
            appl.Visible = true;

            wrkB = appl.Workbooks.Add(Missing.Value);
            wrkS = (Excel.Worksheet)wrkB.Sheets.Item[1];

        }

        public void setup(bool Interactive = true)
        {
            if (!Interactive)
                SetLock(Interactive);
            wrkS.Cells[2, 1].Value = "Sequential";
            wrkS.Cells[3, 1] = "Parallel";
            wrkS.Cells[4, 1] = "Parallel^2";
        }

        public void WriteColumn(int Iteration, TimeSpan Seq, TimeSpan Para, TimeSpan DPara)
        {
            wrkS.Cells[1, DataPoint_StartColumn + Iteration ] = Iteration.ToString();

            //Excel.Range ran = (Excel.Range)wrkS.Range[wrkS.Cells[1, DataPoint_StartColumn + Iteration], wrkS.Cells[4, DataPoint_StartColumn + Iteration]];
            //ran.NumberFormat = "@";

            wrkS.Cells[2, DataPoint_StartColumn + Iteration ] = $"{Seq:ss\\.fffffff}";
            wrkS.Cells[3, DataPoint_StartColumn + Iteration ] = $"{Para:ss\\.fffffff}";
            wrkS.Cells[4, DataPoint_StartColumn + Iteration ] = $"{DPara:ss\\.fffffff}";
        }

        public void SetLock(bool state) => appl.Interactive = state;
    }
}