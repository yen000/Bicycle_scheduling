using System;
using ILOG.Concert;
using ILOG.CPLEX;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using Excel = Microsoft.Office.Interop.Excel;
using Exception = System.Exception;


namespace Bicycle_scheduling
{
    class Program
    {
        const int Machine_type = 4;
        const int Crew_num = 15;
        static void Main(string[] args)
        {
            Console.WriteLine("<Application start>\n");

            # region Declare Parameters

            // required all parameters

            List<string> Job_name = new List<string>();

            List<int> Job_demand = new List<int>();

            List<List<int>> Job_list = new List<List<int>>();

            List<double> Pitch_required = new List<double>();

            List<List<int>> Transportation_time = new List<List<int>>();

            List<List<List<int>>> Process_time = new List<List<List<int>>>();

            List<List<List<int>>> Given_x_variable_solutions = new List<List<List<int>>>();

            List<List<int>> Given_e_variable_solutions = new List<List<int>>();

            # endregion

            #region Read data

            // input files from outsourcing excel

            Microsoft.Office.Interop.Excel.Application excel_reader = new Microsoft.Office.Interop.Excel.Application();
            Workbook reading_workbook = null;
            Worksheet reading_worksheet = null;

            List<string> sheet_names = new List<string>() { "資料表說明", "途程資料", "貼標工時", "限制式", "工單資料表", "車種+配色與物件料號空格定義" };

            //string input_file_name = "巨大塗裝排程資料for清大_20201220_1.xlsx";
            // string input_file_name = Directory.GetCurrentDirectory() + "\\巨大塗裝排程資料for清大_20201220_1.xlsx";
            string input_file_name = Directory.GetCurrentDirectory() + "\\0000.xlsx";
            Console.WriteLine("Read file path:\n" + input_file_name + "\n");

            excel_reader.Workbooks.Open(input_file_name, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            reading_workbook = excel_reader.Workbooks[1];
            reading_workbook.Save();

            // note: sheet index starts from 1

            // get transportation time

            List<List<string>> Sheet1_Data = Read_Fist_Sheet((Worksheet)reading_workbook.Worksheets[1]);

            // get other data from excel files
            List<List<string>> Sheet2_Data = Read_Sheet((Worksheet)reading_workbook.Worksheets[2]);

            List<List<string>> Sheet3_Data = Read_Sheet((Worksheet)reading_workbook.Worksheets[4]);

            List<List<string>> Sheet5_Data = Read_Sheet((Worksheet)reading_workbook.Worksheets[5]);

            reading_workbook.Close();
            excel_reader.Quit();

            #endregion

            # region Data Preprocessing

            Set_Job_Name_And_Job_Demand(Sheet5_Data, Job_name, Job_demand);

            Set_Job_List_And_Pitch_Required(Sheet2_Data, Job_name, Job_demand, Job_list, Pitch_required);


            for (int i = 0; i < Job_name.Count; i++)
            {
                string output = Job_name[i] + "[" + Pitch_required[i].ToString() + "]: ";

                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    output += Job_list[i][j].ToString() + ", ";
                }

                Console.WriteLine(output.TrimEnd(new char[] { ',', ' ' }));
            }


            Set_Transportation_time(Sheet1_Data[1], Transportation_time);

            Set_Process_Time(Sheet1_Data[0], Sheet3_Data, Process_time, Job_list, Pitch_required, Job_name, Job_demand);

            Set_Given_x_variable_solutions(Given_x_variable_solutions, Job_list);

            Set_Given_e_variable_solutions(Given_e_variable_solutions, Job_list);


            #endregion

            #region Mathematical model

            Console.WriteLine("<Cplex starts>");
            int M = 999999;

            int Job_amount = Job_list.Count;
            List<int> machine_k = new List<int> { 0, 1, 2, 3 }; // 0:大 1:小 2:貼 3:金

            Cplex model = new Cplex();

            model.SetParam(Cplex.Param.Emphasis.Memory, true);
            model.SetParam(Cplex.Param.Threads, 4);
            //model.SetParam(Cplex.Param.MIP.Limits.Solutions, 1);
            //model.SetParam(Cplex.Param.TimeLimit, 180);

            // Declare variable
            INumVar[][][] s_varible = new INumVar[Job_amount][][];
            INumVar[][][] c_varible = new INumVar[Job_amount][][];
            INumVar[][][] x_varible = new INumVar[Job_amount][][];
            INumVar[][] e_varible = new INumVar[Job_amount][];
            INumVar[][][][][] y_varible = new INumVar[Job_amount][][][][];
            INumVar Cmax;


            INumVar[][][] A = new INumVar[Job_amount][][];
            for (int i = 0; i < Job_amount; i++)
            {
                A[i] = new INumVar[Job_list[i].Count][];

                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    A[i][j] = model.NumVarArray(machine_k.Count, 0, int.MaxValue, NumVarType.Int);
                }
            }


            INumVar[][][][][] B = new INumVar[Job_amount][][][][];
            for (int i = 0; i < Job_amount; i++)
            {
                B[i] = new INumVar[Job_list[i].Count][][][];
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    B[i][j] = new INumVar[Job_amount][][];

                    for (int i_p = 0; i_p < Job_amount; i_p++)
                    {
                        B[i][j][i_p] = new INumVar[Job_list[i_p].Count][];

                        for (int j_p = 0; j_p < Job_list[i_p].Count; j_p++)
                        {
                            B[i][j][i_p][j_p] = model.NumVarArray(machine_k.Count, 0, int.MaxValue, NumVarType.Int);
                        }
                    }
                }
            }

            INumVar[][] C = new INumVar[Job_amount][];
            for (int i = 0; i < Job_amount; i++)
            {
                C[i] = model.NumVarArray(Job_list[i].Count, 0, int.MaxValue, NumVarType.Int);
            }


            for (int i = 0; i < Job_amount; i++)
            {
                s_varible[i] = new INumVar[Job_list[i].Count][];

                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    s_varible[i][j] = model.NumVarArray(machine_k.Count, 0, float.MaxValue, NumVarType.Float);
                }
            }

            for (int i = 0; i < Job_amount; i++)
            {
                c_varible[i] = new INumVar[Job_list[i].Count][];

                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    c_varible[i][j] = model.NumVarArray(machine_k.Count, 0, float.MaxValue, NumVarType.Float);
                }
            }

            for (int i = 0; i < Job_amount; i++)
            {
                x_varible[i] = new INumVar[Job_list[i].Count][];

                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    x_varible[i][j] = model.NumVarArray(machine_k.Count, 0, int.MaxValue, NumVarType.Bool);
                }
            }

            for (int i = 0; i < Job_amount; i++)
            {
                e_varible[i] = model.NumVarArray(Job_list[i].Count, 0, int.MaxValue, NumVarType.Bool);
            }

            for (int i = 0; i < Job_amount; i++)
            {
                y_varible[i] = new INumVar[Job_list[i].Count][][][];
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    y_varible[i][j] = new INumVar[Job_amount][][];

                    for (int i_p = 0; i_p < Job_amount; i_p++)
                    {
                        y_varible[i][j][i_p] = new INumVar[Job_list[i_p].Count][];

                        for (int j_p = 0; j_p < Job_list[i_p].Count; j_p++)
                        {
                            y_varible[i][j][i_p][j_p] = model.NumVarArray(machine_k.Count, 0, int.MaxValue, NumVarType.Bool);
                        }
                    }
                }
            }


            Cmax = model.NumVar(0, float.MaxValue);


            //Objective function
            model.AddMinimize(Cmax);


            //Constraint 1
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    for (int k = 0; k < 4; k++)
                    {
                        ILinearNumExpr constraint1 = model.LinearNumExpr();
                        constraint1.AddTerm(1, c_varible[i][j][k]);
                        constraint1.AddTerm(-1, s_varible[i][j][k]);
                        model.AddEq(constraint1, Process_time[i][j][k]);
                    }

                }
            }

            //Constraint 2
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    for (int k = 0; k < machine_k.Count; k++)
                    {
                        ILinearNumExpr constraint2 = model.LinearNumExpr();
                        constraint2.AddTerm(1, Cmax);
                        constraint2.AddTerm(-1, c_varible[i][j][k]);
                        model.AddGe(constraint2, 0);
                    }
                }
            }

            //Constraint A
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 1; j < Job_list[i].Count; j++)
                {
                    for (int k = 0; k < 4; k++)
                    {
                        ILinearNumExpr constraint_A = model.LinearNumExpr();
                        if (j - 2 >= 0)
                        {
                            constraint_A.AddTerm(-1, e_varible[i][j - 2]);
                        }
                        if (j - 1 >= 0)
                        {
                            constraint_A.AddTerm(-1, e_varible[i][j - 1]);
                        }
                        constraint_A.AddTerm(1, x_varible[i][j - 1][k]);
                        constraint_A.AddTerm(1, A[i][j - 1][k]);

                        model.AddEq(constraint_A, 1);
                    }
                }
            }

            //Constraint 3 4
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 1; j < Job_list[i].Count; j++)
                {
                    for (int k = 0; k < 4; k++)
                    {
                        for (int k_p = 0; k_p < 4; k_p++)
                        {
                            ILinearNumExpr constraint3 = model.LinearNumExpr();
                            constraint3.AddTerm(1, s_varible[i][j][k_p]);
                            constraint3.AddTerm(-1, c_varible[i][j - 1][k]);
                            constraint3.AddTerm(M, A[i][j - 1][k]);

                            int cost = 0;
                            if (k_p == 0 || k_p == 1)
                            {
                                cost = 138 + Transportation_time[k][k_p];
                            }
                            /*else if(k_p==3)
                            {
                                cost = 120 + Transportation_time[k][k_p];
                            }*/
                            else
                            {
                                cost = Transportation_time[k][k_p];
                            }


                            // model.AddGe(constraint3, cost);

                            model.Add(model.IfThen(model.Eq(A[i][j - 1][k], 0.0), model.AddGe(constraint3, cost)));

                            //model.AddGe(constraint3, cost);
                        }
                    }
                }
            }

            //Constraint B
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    for (int i_p = 0; i_p < Job_amount; i_p++)
                    {
                        for (int j_p = 0; j_p < Job_list[i_p].Count; j_p++)
                        {
                            for (int k = 0; k < 4; k++)
                            {
                                ILinearNumExpr constraint_B = model.LinearNumExpr();
                                constraint_B.AddTerm(1, x_varible[i][j][k]);
                                constraint_B.AddTerm(1, x_varible[i_p][j_p][k]);
                                constraint_B.AddTerm(-1, y_varible[i][j][i_p][j_p][k]);
                                constraint_B.AddTerm(1, B[i][j][i_p][j_p][k]);
                                model.AddEq(constraint_B, 2);
                            }
                        }
                    }
                }
            }




            //Constraint 5 - 1
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    for (int i_p = 0; i_p < Job_amount; i_p++)
                    {
                        for (int j_p = 0; j_p < Job_list[i_p].Count; j_p++)
                        {
                            for (int k = 0; k < 2; k++)
                            {
                                if (i != i_p || j != j_p)
                                {
                                    ILinearNumExpr constraint5_1 = model.LinearNumExpr();
                                    constraint5_1.AddTerm(1, s_varible[i][j][k]);
                                    constraint5_1.AddTerm(-1, s_varible[i_p][j_p][k]);
                                    constraint5_1.AddTerm(M, B[i][j][i_p][j_p][k]);

                                    //model.AddGe(constraint5_1, 23 * Pitch_required[i_p] + 138 - 2 * M);
                                    model.Add(model.IfThen(model.Eq(B[i][j][i_p][j_p][k], 0.0), model.AddGe(constraint5_1, 23 * Pitch_required[i_p] + 138)));

                                }
                            }
                        }
                    }
                }
            }

            //Constraint 5 - 2
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    for (int i_p = 0; i_p < Job_amount; i_p++)
                    {
                        for (int j_p = 0; j_p < Job_list[i_p].Count; j_p++)
                        {
                            for (int k = 2; k < 3; k++)
                            {
                                if (i != i_p || j != j_p)
                                {
                                    ILinearNumExpr constraint5_2 = model.LinearNumExpr();
                                    constraint5_2.AddTerm(1, s_varible[i][j][k]);
                                    constraint5_2.AddTerm(-1, s_varible[i_p][j_p][k]);
                                    constraint5_2.AddTerm(M, B[i][j][i_p][j_p][k]);
                                    // model.AddGe(constraint5_2, 0 - 2 * M);
                                    model.Add(model.IfThen(model.Eq(B[i][j][i_p][j_p][k], 0.0), model.AddGe(constraint5_2, 0)));
                                }
                            }
                        }
                    }
                }
            }

            //Constraint 5 - 3
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    for (int i_p = 0; i_p < Job_amount; i_p++)
                    {
                        for (int j_p = 0; j_p < Job_list[i_p].Count; j_p++)
                        {
                            for (int k = 3; k < 4; k++)
                            {
                                if (i != i_p || j != j_p)
                                {
                                    ILinearNumExpr constraint5_3 = model.LinearNumExpr();
                                    constraint5_3.AddTerm(1, s_varible[i][j][k]);
                                    constraint5_3.AddTerm(-1, s_varible[i_p][j_p][k]);
                                    constraint5_3.AddTerm(M, B[i][j][i_p][j_p][k]);

                                    // model.AddGe(constraint5_3, 20 * Pitch_required[i_p] + 0 - 2 * M);
                                    model.Add(model.IfThen(model.Eq(B[i][j][i_p][j_p][k], 0.0), model.AddGe(constraint5_3, 20 * Pitch_required[i_p])));
                                }
                            }
                        }
                    }
                }
            }
            /*
            // Constraint 6 - 1
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    for (int i_p = 0; i_p < Job_amount; i_p++)
                    {
                        for (int j_p = 0; j_p < Job_list[i_p].Count; j_p++)
                        {
                            for (int k = 0; k < 2; k++)
                            {
                                if (i != i_p || j != j_p)
                                {
                                    ILinearNumExpr constraint6_1 = model.LinearNumExpr();
                                    constraint6_1.AddTerm(-1, s_varible[i][j][k]);
                                    constraint6_1.AddTerm(1, s_varible[i_p][j_p][k]);
                                    constraint6_1.AddTerm(-M, x_varible[i][j][k]);
                                    constraint6_1.AddTerm(-M, x_varible[i_p][j_p][k]);
                                    constraint6_1.AddTerm(M, y_varible[i_p][j_p][i][j][k]);
                                    model.AddGe(constraint6_1, 23 * Pitch_required[i] + 138 - 2 * M);
                                }
                            }
                        }
                    }
                }
            }

            // Constraint 6 - 2
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    for (int i_p = 0; i_p < Job_amount; i_p++)
                    {
                        for (int j_p = 0; j_p < Job_list[i_p].Count; j_p++)
                        {
                            for (int k = 2; k < 3; k++)
                            {
                                if (i != i_p || j != j_p)
                                {
                                    ILinearNumExpr constraint6_2 = model.LinearNumExpr();
                                    constraint6_2.AddTerm(-1, s_varible[i][j][k]);
                                    constraint6_2.AddTerm(1, s_varible[i_p][j_p][k]);
                                    constraint6_2.AddTerm(-M, x_varible[i][j][k]);
                                    constraint6_2.AddTerm(-M, x_varible[i_p][j_p][k]);
                                    constraint6_2.AddTerm(M, y_varible[i_p][j_p][i][j][k]);
                                    model.AddGe(constraint6_2, -2 * M);
                                }
                            }
                        }
                    }
                }
            }

            // Constraint 6 - 3
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    for (int i_p = 0; i_p < Job_amount; i_p++)
                    {
                        for (int j_p = 0; j_p < Job_list[i_p].Count; j_p++)
                        {
                            for (int k = 3; k < 4; k++)
                            {
                                if (i != i_p || j != j_p)
                                {
                                    ILinearNumExpr constraint6_3 = model.LinearNumExpr();
                                    constraint6_3.AddTerm(-1, s_varible[i][j][k]);
                                    constraint6_3.AddTerm(1, s_varible[i_p][j_p][k]);
                                    constraint6_3.AddTerm(-M, x_varible[i][j][k]);
                                    constraint6_3.AddTerm(-M, x_varible[i_p][j_p][k]);
                                    constraint6_3.AddTerm(M, y_varible[i_p][j_p][i][j][k]);
                                    model.AddGe(constraint6_3, 20 * Pitch_required[i] - 2 * M);
                                }
                            }
                        }
                    }
                }
            }
            */
            // Constraint 7
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    for (int i_p = 0; i_p < Job_amount; i_p++)
                    {
                        for (int j_p = 0; j_p < Job_list[i_p].Count && i != i_p; j_p++)
                        {
                            for (int k = 0; k < machine_k.Count; k++)
                            {
                                ILinearNumExpr constraint7 = model.LinearNumExpr();
                                constraint7.AddTerm(1, y_varible[i_p][j_p][i][j][k]);
                                constraint7.AddTerm(1, y_varible[i][j][i_p][j_p][k]);
                                model.AddEq(constraint7, 1);
                            }
                        }
                    }
                }
            }

            // Constraint 9
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    ILinearNumExpr constraint9 = model.LinearNumExpr();

                    for (int k = 0; k < machine_k.Count; k++)
                    {
                        constraint9.AddTerm(1, x_varible[i][j][k]);
                    }
                    model.AddEq(constraint9, 1);
                }
            }


            //Constraint 10
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    if (j == 0)
                    {
                        ILinearNumExpr constraint10 = model.LinearNumExpr();

                        constraint10.AddTerm(1, e_varible[i][j]);
                        constraint10.AddTerm(-1, x_varible[i][j][0]);

                        model.AddEq(constraint10, 0);
                    }
                    else if (j == 1)
                    {
                        ILinearNumExpr constraint10 = model.LinearNumExpr();

                        constraint10.AddTerm(1, e_varible[i][j - 1]);
                        constraint10.AddTerm(1, e_varible[i][j]);
                        constraint10.AddTerm(-1, x_varible[i][j][0]);

                        model.AddEq(constraint10, 0);
                    }
                    else if (j >= 2)
                    {
                        ILinearNumExpr constraint10 = model.LinearNumExpr();

                        constraint10.AddTerm(1, e_varible[i][j - 2]);
                        constraint10.AddTerm(1, e_varible[i][j - 1]);
                        constraint10.AddTerm(1, e_varible[i][j]);
                        constraint10.AddTerm(-1, x_varible[i][j][0]);

                        model.AddEq(constraint10, 0);
                    }
                }
            }

            /*
            // Constraint 11
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 0; j < Job_list[i].Count-2; j++)
                {
                    ILinearNumExpr constraint11 = model.LinearNumExpr();

                    constraint11.AddTerm(1, e_varible[i][j + 2]);
                    constraint11.AddTerm(1, e_varible[i][j + 1]);
                    constraint11.AddTerm(1, e_varible[i][j]);

                    model.AddLe(constraint11, 1);

                }
            }
            */
            // Constraint Extra(12)
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    for (int k = 0; k < machine_k.Count; k++)
                    {
                        if (Given_x_variable_solutions[i][j][k] != -1)
                        {
                            ILinearNumExpr constraint12 = model.LinearNumExpr();
                            constraint12.AddTerm(1, x_varible[i][j][k]);
                            model.AddEq(constraint12, Given_x_variable_solutions[i][j][k]);
                        }
                    }

                }
            }

            // Constraint Extra(13)
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    if (Given_e_variable_solutions[i][j] != -1)
                    {
                        ILinearNumExpr constraint13 = model.LinearNumExpr();
                        constraint13.AddTerm(1, e_varible[i][j]);
                        model.AddEq(constraint13, Given_e_variable_solutions[i][j]);
                    }
                }
            }

            // Constraint C
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 1; j < Job_list[i].Count; j++)
                {

                    ILinearNumExpr constraint_C = model.LinearNumExpr();

                    if (j - 2 >= 0)
                    {
                        constraint_C.AddTerm(1, e_varible[i][j - 2]);
                    }

                    constraint_C.AddTerm(1, e_varible[i][j - 1]);
                    constraint_C.AddTerm(1, C[i][j - 1]);

                    model.AddEq(constraint_C, 1);

                }
            }


            // Constraint 14
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 1; j < Job_list[i].Count; j++)
                {

                    ILinearNumExpr constraint14 = model.LinearNumExpr();
                    constraint14.AddTerm(1, s_varible[i][j - 1][0]);
                    constraint14.AddTerm(-1, s_varible[i][j][0]);
                    constraint14.AddTerm(M, C[i][j - 1]);

                    // model.AddGe(constraint14, -M);
                    model.Add(model.IfThen(model.Eq(C[i][j - 1], 0.0), model.AddGe(constraint14, 0)));

                }
            }

            // Constraint 14
            for (int i = 0; i < Job_amount; i++)
            {
                for (int j = 1; j < Job_list[i].Count; j++)
                {

                    ILinearNumExpr constraint14 = model.LinearNumExpr();
                    constraint14.AddTerm(-1, s_varible[i][j - 1][0]);
                    constraint14.AddTerm(1, s_varible[i][j][0]);


                    model.AddGe(constraint14, 0);

                }
            }

            model.Solve();

            // Objective value result
            Console.WriteLine("Objective value: " + model.GetObjValue());

            // Result
            double[][][] s_result = new double[Job_amount][][];
            for (int i = 0; i < Job_amount; i++)
            {
                s_result[i] = new double[Job_list[i].Count][];

                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    s_result[i][j] = new double[Machine_type];

                    s_result[i][j] = model.GetValues(s_varible[i][j]);

                }
            }

            double[][][] c_result = new double[Job_amount][][];
            for (int i = 0; i < Job_amount; i++)
            {
                c_result[i] = new double[Job_list[i].Count][];

                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    c_result[i][j] = new double[Machine_type];

                    c_result[i][j] = model.GetValues(c_varible[i][j]);
                }
            }

            double[][][] x_result = new double[Job_amount][][];
            for (int i = 0; i < Job_amount; i++)
            {
                x_result[i] = new double[Job_list[i].Count][];

                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    x_result[i][j] = new double[Machine_type];

                    x_result[i][j] = model.GetValues(x_varible[i][j]);
                }
            }

            double[][] e_result = new double[Job_amount][];

            for (int i = 0; i < Job_amount; i++)
            {
                e_result[i] = model.GetValues(e_varible[i]);
            }

            for (int i = 0; i < Job_list.Count; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    for (int k = 0; k < Machine_type; k++)
                    {
                        if (Math.Round(x_result[i][j][k]) == 1)
                        {

                            string output = "";

                            output += "Machine #" + k.ToString() + ", ";
                            output += "x[" + i.ToString() + "," + j.ToString() + "," + k.ToString() + "]= " + Math.Round(x_result[i][j][k]).ToString() + ",     ";
                            //output += "e[" + i.ToString() + "," + j.ToString() + "]= " + Math.Round(e_result[i][j]).ToString() + ",     ";
                            output += "s[" + i.ToString() + "," + j.ToString() + "," + k.ToString() + "]= " + Math.Round(s_result[i][j][k]).ToString() + ",     ";
                            output += "c[" + i.ToString() + "," + j.ToString() + "," + k.ToString() + "]= " + Math.Round(c_result[i][j][k]).ToString();

                            Console.WriteLine(output);

                        }
                    }
                }
                Console.WriteLine("----------");
            }
            /*
            Console.WriteLine("---");

            Console.WriteLine("x[1.0.0]= " + x_result[1][0][0]);
            Console.WriteLine("x[1.0.1]= " + x_result[1][0][1]);
            Console.WriteLine("x[1.0.2]= " + x_result[1][0][2]);
            Console.WriteLine("x[1.0.3]= " + x_result[1][0][3]);

            Console.WriteLine("---");

            Console.WriteLine("eij-1= " + e_result[1][0]);
            Console.WriteLine("xij-1k= " + x_result[1][0][1]);


            Console.WriteLine("---");
            
            for (int k = 0; k < Machine_type; k++)
            {
                Console.WriteLine("s[1,1,1] >= " + c_result[1][1][k] + " + trans(" + Transportation_time[k][1] + ") +138 -M*( 1 - " + x_result[1][0][k] + ")");
            }
            */
            //Console.WriteLine("Y03200= " + model.GetValue(y_varible[0][3][2][0][0]));
            //Console.WriteLine("Y20030= " + model.GetValue(y_varible[2][0][0][3][0]));


            //Console.WriteLine("pithch needed: " + (23 * Pitch_required[0] + 138).ToString());

            //display constraint of 5&6
            /*
            for (int i = 0; i < Job_list.Count; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    for (int i_pi = 0; i_pi < Job_list.Count && i != i_pi; i_pi++)
                    {
                        for (int j_pi = 0; j_pi < Job_list[i_pi].Count; j_pi++)
                        {
                            for (int k = 1; k < 2; k++)
                            {
                                if (x_result[i][j][k] == 1 && x_result[i_pi][j_pi][k] == 1 && model.GetValue(y_varible[i][j][i_pi][j_pi][k]) == 0)
                                {
                                    Console.WriteLine("S[" + i.ToString() + "," + j.ToString() + "," + k.ToString() + "] >= " + "S'[" + i_pi.ToString() + "," + j_pi.ToString() + "," + k.ToString() + "] +  " + (23 * Pitch_required[i_pi] + 138).ToString());
                                }
                                if (x_result[i][j][k] == 1 && x_result[i_pi][j_pi][k] == 1 && model.GetValue(y_varible[i_pi][j_pi][i][j][k]) == 0)
                                {
                                    Console.WriteLine("S'[" + i_pi.ToString() + "," + j_pi.ToString() + "," + k.ToString() + "] >= " + "S[" + i.ToString() + "," + j.ToString() + "," + k.ToString() + "] +  " + (23 * Pitch_required[i] + 138).ToString());
                                }
                            }
                        }
                    }
                }
            }
            */
            /*
            for(int k = 0; k < Machine_type; k++)
            {
                Console.WriteLine("s[10,3,1] >= " + c_result[10][2][k] + " + trans(" + Transportation_time[k][1] + ") +138 -M*( " + e_result[10][1] + " + " + e_result[10][2] + ")");
            }
            */
            /*
            for (int i = 0; i < Job_list.Count; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    for (int i_pi = 0; i_pi < Job_list.Count && i != i_pi; i_pi++)
                    {
                        for (int j_pi = 0; j_pi < Job_list[i_pi].Count; j_pi++)
                        {
                            for (int k = 0; k < 1; k++)
                            {
                                if (x_result[i][j][k] == 1 && x_result[i_pi][j_pi][k] == 1)
                                {
                                    Console.WriteLine(model.GetValue(y_varible[i][j][i_pi][j_pi][k]) + "," + (model.GetValue(y_varible[i_pi][j_pi][i][j][k])));
                                }
                            }
                        }
                    }
                }
            }
            */
            /*
            Console.WriteLine(Process_time[2][0][0]);

            */
            double Cmax_result = model.GetValue(Cmax);
            #endregion

            #region output file

            // 設定儲存檔名，不用設定副檔名，系統自動判斷 excel 版本，產生 .xls 或 .xlsx 副檔名
            string pathFile = Directory.GetCurrentDirectory() + "\\result.xlsx";
            //string pathFile = @"D:\test";

            Excel.Application excelApp;
            Excel._Workbook wBook;
            Excel._Worksheet wSheet;
            Range wRange;

            // 開啟一個新的應用程式
            excelApp = new Excel.Application();

            // 讓Excel文件可見
            excelApp.Visible = true;

            // 停用警告訊息
            excelApp.DisplayAlerts = false;

            // 加入新的活頁簿
            excelApp.Workbooks.Add(Type.Missing);

            // 引用第一個活頁簿
            wBook = excelApp.Workbooks[1];

            // 設定活頁簿焦點
            wBook.Activate();

            try
            {
                // 引用第一個工作表

                excelApp.Worksheets.Add();
                wSheet = (Excel._Worksheet)wBook.Worksheets[1];

                // 命名工作表的名稱
                wSheet.Name = "物件程序";

                // 設定工作表焦點
                wSheet.Activate();

                //  excelApp.Cells[1, 1] = "Excel測試";

                //  int i = 5;
                // 設定第1列資料

                // 設定第1列顏色
                //  wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 1+i]];
                //  wRange.Select();
                /*
                //  wRange.Font.Color = ColorTranslator.ToOle(Color.White);
                // wRange.Interior.Color = ColorTranslator.ToOle(Color.DimGray);



                 }*/

                //int[,,] start_point = new int[10000, 10000, 4];
                //int[,,] x_output = new int[10000, 10000, 4];

                double[][][] start_point = s_result;

                double[][][] x_output = x_result;

                int index = 100000;  //index 為訂單數量

                //  int Job_list[i].Count = 100000; //Job_list[i].Count 為狀態


                List<int> add_order = new List<int>(); //決定開始順序, i
                List<int> add_place = new List<int>(); //決定加入位置, k
                List<double> add_time = new List<double>();  //加入時間大小, s

                //決定一開始加入大線小線順序
                for (int i = 0; i < Job_list.Count; i++)
                {
                    for (int j = 0; j < Job_list[i].Count; j++)
                    {
                        for (int k = 0; k < Machine_type; k++)
                        {
                            if ((j == 0 && start_point[i][j][k] == 0))
                            {
                                if ((k == 0 || k == 1) && x_output[i][j][k] == 1)
                                {
                                    add_time.Add(start_point[i][j][k]);                 //存入最早開始時間
                                    add_order.Add(i);                               //存入訂單號

                                }
                            }
                            if (j == 0 && start_point[i][j][k] != 0)
                            {
                                if ((k == 0 || k == 1) && x_output[i][j][k] == 1)
                                {
                                    add_time.Add(start_point[i][j][k]);
                                    add_order.Add(i);
                                    //  add_place.Add(k);

                                }
                            }
                        }
                    }
                }

                //add_time.Sort(add_time,add_order);  //時間做排序 最小到最大


                int count_0 = 0;
                int count_1 = 0;


                double[] ad_time = add_time.ToArray();
                // double[] ad_time1 = add_time.ToArray();
                int[] ad_order = add_order.ToArray();
                int[] ad_place = new int[add_time.Count];         //    new int[ad_time.Length];
                Array.Sort(ad_time, ad_order);
                //  Array.Sort(ad_time1, ad_place);
                //時間做排序 最小到最大
                //將大線小線位置存入 以及物件料號


                for (int i = 0; i < Job_list.Count; i++)
                {
                    for (int j = 0; j < Job_list[i].Count; j++)
                    {
                        for (int k = 0; k < Machine_type; k++)
                        {
                            for (int l = 0; l < ad_time.Length; l++)
                            {
                                if (start_point[i][j][k] == ad_time[l] && x_output[i][j][k] == 1 && ad_order[l] == i)
                                {
                                    // ad_order[l] = i;
                                    ad_place[l] = k;
                                }
                                //add_time.Add(start_point[i, j, k]);
                            }
                        }
                    }
                }
                //依照順序印出加入時間順序
                excelApp.Cells[1, 1] = "物件料號";
                excelApp.Cells[1, 2] = "開始時間";
                excelApp.Cells[1, 3] = "開始位置";
                excelApp.Cells[1, 4] = "開始位置順序";
                excelApp.Cells[1, 5] = "組裝過程";


                int[] order = new int[ad_time.Length];
                int count = 0;
                int location = 0;
                for (int i = 0; i < ad_time.Length; i++)
                {
                    //第一個步驟
                    if (ad_place[i] == 0)
                    {
                        count_0 = count_0 + 1;
                        order[i] = count_0;
                        excelApp.Cells[2 + i, 1] = string.Format("{0}", Job_name[ad_order[i]]);
                        excelApp.Cells[2 + i, 2] = string.Format("{0}", ad_time[i]);
                        excelApp.Cells[2 + i, 3] = "大線";
                        excelApp.Cells[2 + i, 4] = string.Format("{0}", order[i]);
                        excelApp.Cells[2 + i, 5] = "大線";
                        count = 1;
                        /*       for (int ii = 0; ii < ad_order.Length; ii++)
                                {
                                    for (int j = 1; j < Job_list[ad_order[ii]].Count-1; j++)
                                    {
                                        if (x_output[ad_order[ii]][j][0] == 1 && start_point[ad_order[ii]][j][0] < start_point[ad_order[i]][0][0] && start_point[ad_order[ii]][j][1] != start_point[ad_order[ii]][j + 1][1])
                                        {
                                            count_1 = count_1 + 1;
                                            order[i] = count_1;
                                        }
                                    }
                                }*/
                        excelApp.Cells[2 + i, 4] = string.Format("{0}", order[i]);


                        for (int j = 0; j < Job_list[ad_order[i]].Count - 1; j++)
                        {
                            for (int k = 0; k < Machine_type; k++)
                            {
                                if (x_output[ad_order[i]][j][k] == 0 && x_output[ad_order[i]][j + 1][k] == 1)
                                {
                                    if (k == 0)
                                    {
                                        excelApp.Cells[2 + i, 5 + count] = "大線";
                                        count++;
                                        for (int j1 = j + 1; j1 < Job_list[ad_order[i]].Count - 2; j1++)
                                        {
                                            if (x_output[ad_order[i]][j1 - 1][0] == 1 && x_output[ad_order[i]][j1][0] == 1 && x_output[ad_order[i]][j1 + 1][0] == 1 && x_output[ad_order[i]][j1 + 2][0] == 1)
                                            {
                                                excelApp.Cells[2 + i, 5 + count] = "大線";
                                                count++;
                                            }
                                        }
                                    }
                                    if (k == 1)
                                    {
                                        excelApp.Cells[2 + i, 5 + count] = "小線";
                                        count++;
                                    }
                                    if (k == 2)
                                    {
                                        excelApp.Cells[2 + i, 5 + count] = "貼標";
                                        count++;
                                    }
                                    if (k == 3)
                                    {
                                        excelApp.Cells[2 + i, 5 + count] = "金油線";
                                        count++;
                                    }
                                }
                                if (x_output[ad_order[i]][j][k] == 1 && x_output[ad_order[i]][j + 1][k] == 1)
                                {
                                    if (k == 1)
                                    {
                                        excelApp.Cells[2 + i, 5 + count] = "小線";
                                        count++;
                                    }
                                    if (k == 3)
                                    {
                                        excelApp.Cells[2 + i, 5 + count] = "金油線";
                                        count++;
                                    }
                                }
                            }
                        }
                    }
                    if (ad_place[i] == 1)
                    {
                        count_1 = count_1 + 1;
                        order[i] = count_1;
                        excelApp.Cells[2 + i, 1] = string.Format("{0}", Job_name[ad_order[i]]);
                        excelApp.Cells[2 + i, 2] = string.Format("{0}", ad_time[i]);
                        excelApp.Cells[2 + i, 3] = "小線";
                        excelApp.Cells[2 + i, 4] = string.Format("{0}", order[i]);
                        excelApp.Cells[2 + i, 5] = "小線";
                        count = 1;
                        /*  for (int ii = 0; ii < ad_order.Length; ii++)
                            {
                                for (int j = 1; j < Job_list[ad_order[ii]].Count-1; j++)
                                {
                                    if (x_output[ad_order[ii]][j][1] == 1 && start_point[ad_order[ii]][j][1] < start_point[ad_order[i]][0][1] && start_point[ad_order[ii]][j][1]!= start_point[ad_order[ii]][j+1][1])
                                    {
                                        count_1 = count_1 + 1;
                                        order[i] = count_1;
                                    }
                                }
                            }*/
                        excelApp.Cells[2 + i, 4] = string.Format("{0}", order[i]);


                        for (int j = 0; j < Job_list[ad_order[i]].Count - 1; j++)
                        {
                            for (int k = 0; k < Machine_type; k++)
                            {
                                if (x_output[ad_order[i]][j][k] == 1 && x_output[ad_order[i]][j + 1][k] == 1)
                                {
                                    if (k == 1)
                                    {
                                        excelApp.Cells[2 + i, 5 + count] = "小線";
                                        count++;
                                    }
                                    if (k == 3)
                                    {
                                        excelApp.Cells[2 + i, 5 + count] = "金油線";
                                        count++;
                                    }
                                }
                                if (x_output[ad_order[i]][j][k] == 0 && x_output[ad_order[i]][j + 1][k] == 1)
                                {
                                    if (k == 0)
                                    {
                                        excelApp.Cells[2 + i, 5 + count] = "大線";
                                        count++;
                                        for (int j1 = j + 1; j1 < Job_list[ad_order[i]].Count - 2; j1++)
                                        {
                                            if (x_output[ad_order[i]][j1 - 1][0] == 1 && x_output[ad_order[i]][j1][0] == 1 && x_output[ad_order[i]][j1 + 1][0] == 1 && x_output[ad_order[i]][j1 + 2][0] == 1)
                                            {
                                                excelApp.Cells[2 + i, 5 + count] = "大線";
                                                count++;
                                            }
                                        }
                                    }
                                    if (k == 1)
                                    {
                                        excelApp.Cells[2 + i, 5 + count] = "小線";
                                        count++;
                                    }
                                    if (k == 2)
                                    {
                                        excelApp.Cells[2 + i, 5 + count] = "貼標";
                                        count++;
                                    }
                                    if (k == 3)
                                    {
                                        excelApp.Cells[2 + i, 5 + count] = "金油線";

                                        count++;
                                    }
                                }
                            }
                        }
                    }

                    location = 2 + i;
                }
                wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[100, 100]];
                wRange.Select();
                wRange.Columns.AutoFit();

                wSheet = (Excel._Worksheet)wBook.Worksheets[2];
                Excel._Worksheet wSheet1 = (Excel._Worksheet)wBook.Worksheets[1];

                // 命名工作表的名稱
                wSheet.Name = "各線流程";

                // 設定工作表焦點
                wSheet.Activate();

                //工作站作業順序
                int location2 = 1;
                location = location + 5;
                excelApp.Cells[1, location2] = "工作站作業順序";
                excelApp.Cells[2, location2] = "大線";
                excelApp.Cells[2, location2 + 1] = "小線";
                excelApp.Cells[2, location2 + 2] = "貼標";
                excelApp.Cells[2, location2 + 3] = "金油線";
                int location1 = location;
                List<double> work_at_0 = new List<double>(); //決定加入位置, k
                List<double> work_at_0_job = new List<double>();
                List<double> work_at_1 = new List<double>();  //加入時間大小, s
                List<double> work_at_1_job = new List<double>();
                List<double> work_at_2 = new List<double>();
                List<double> work_at_2_job = new List<double>();
                List<double> work_at_3 = new List<double>();
                List<double> work_at_3_job = new List<double>();


                for (int i = 0; i < Job_list.Count; i++)
                {
                    for (int j = 0; j < Job_list[i].Count - 1; j++)
                    {
                        for (int k = 0; k < Machine_type; k++)
                        {
                            if (x_output[i][j][k] == 1)
                            {
                                if (k == 0 && x_output[i][j + 1][k] == 0 && x_output[i][j][k] == 1)
                                {
                                    work_at_0.Add(start_point[i][j][0]);
                                    work_at_0_job.Add(i);
                                }
                                if (k == 1)
                                {
                                    work_at_1.Add(start_point[i][j][1]);
                                    work_at_1_job.Add(i);
                                }
                                if (k == 2)
                                {
                                    work_at_2.Add(start_point[i][j][2]);
                                    work_at_2_job.Add(i);
                                }

                            }
                        }
                    }
                }

                for (int i = 0; i < Job_list.Count; i++)
                {
                    for (int j = 2; j < Job_list[i].Count; j++)
                    {
                        for (int k = 0; k < Machine_type; k++)
                        {
                            if (k == 3 && x_output[i][j][3] == 1)
                            {
                                work_at_3.Add(start_point[i][j][3]);
                                work_at_3_job.Add(i);
                            }
                        }
                    }
                }



                double[] work_at_0_ = work_at_0.ToArray();
                double[] work_at_0_job_ = work_at_0_job.ToArray();
                double[] work_at_1_ = work_at_1.ToArray();
                double[] work_at_1_job_ = work_at_1_job.ToArray();
                double[] work_at_2_ = work_at_2.ToArray();
                double[] work_at_2_job_ = work_at_2_job.ToArray();
                double[] work_at_3_ = work_at_3.ToArray();
                double[] work_at_3_job_ = work_at_3_job.ToArray();

                Array.Sort(work_at_0_, work_at_0_job_);
                Array.Sort(work_at_1_, work_at_1_job_);
                Array.Sort(work_at_2_, work_at_2_job_);
                Array.Sort(work_at_3_, work_at_3_job_);

                for (int i = 0; i < work_at_0_job_.Length; i++)
                {
                    excelApp.Cells[3 + i, location2] = string.Format("{0}", Job_name[(int)work_at_0_job_[i]]); ;
                }
                for (int i = 0; i < work_at_1_job_.Length; i++)
                {
                    excelApp.Cells[3 + i, location2 + 1] = string.Format("{0}", Job_name[(int)work_at_1_job_[i]]); ;
                }
                for (int i = 0; i < work_at_2_job_.Length; i++)
                {
                    excelApp.Cells[3 + i, location2 + 2] = string.Format("{0}", Job_name[(int)work_at_2_job_[i]]); ;
                }
                for (int i = 0; i < work_at_3_job_.Length; i++)
                {
                    excelApp.Cells[3 + i, location2 + 3] = string.Format("{0}", Job_name[(int)work_at_3_job_[i]]); ;
                }
                int min = 0;

                int[] oorder = new int[10000];
                count = 0;

                for (int i = 0; i < ad_time.Length; i++) // <----------------------------------------------
                {
                    if (ad_place[i] == 0)
                    {
                        for (int j = 0; j < work_at_0_job_.Length; j++)
                        {
                            if (work_at_0_job_[j] == ad_order[i])
                            {
                                oorder[count] = j + 1;
                                count = count + 1;
                            }
                        }

                        min = oorder[0];

                        for (int j = 0; j < work_at_0_job_.Length; j++)
                        {

                            if (min > oorder[j] && oorder[j] != 0)
                            {
                                min = oorder[j];
                            }
                        }
                        order[i] = min;
                        wSheet1.Cells[2 + i, 4] = string.Format("{0}", order[i]);


                    }
                    if (ad_place[i] == 1)
                    {
                        for (int j = 0; j < work_at_1_job_.Length; j++)
                        {
                            if (work_at_1_job_[j] == ad_order[i])
                            {
                                oorder[count] = j + 1;
                                count = count + 1;
                            }
                        }
                        min = oorder[0];
                        for (int j = 0; j < work_at_1_job_.Length; j++)
                        {

                            if (min > oorder[j] && oorder[j] != 0)
                            {
                                min = oorder[j];
                            }
                        }
                        order[i] = min;
                        wSheet1.Cells[2 + i, 4] = string.Format("{0}", order[i]);

                    }
                    for (int j = 0; j < oorder.Length; j++)
                    {

                        oorder[j] = 0;
                    }
                    count = 0;

                }
                /*
                location2 = location2 + 10;
                for(int i=0;i<Job_name.Count;i++)
                {
                    excelApp.Cells[location2+i, 1] = string.Format("{0}", Job_name[i]);
                    excelApp.Cells[location2 + i, 2] = string.Format("{0}", i);
                }*/
                //  for(int o)

                //自動調整欄寬
                wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[100, 100]];
                wRange.Select();
                wRange.Columns.AutoFit();








                /*     for(int i=0;i<ad_time.Length;i++)
                     {

                             if (ad_place[i] == 0)
                             {
                                 for (int ii = 0; ii < work_at_0_job_.Length; ii++)
                                 { 
                                 if(work_at_0_job_[ii]==ad_order[i])
                                 {
                                     order[i] = ii + 1;
                                     excelApp.Cells[2 + i, 4] = string.Format("{0}", 0);
                                 }


                                 }

                                        // excelApp.Cells[2 + i, 4] = string.Format("{0}", 0);

                             }
                             if (ad_place[i] == 1)
                             {

                             for (int ii = 0; ii < work_at_1_job_.Length; ii++)
                             {
                                 if (work_at_1_job_[ii] == ad_order[i])
                                 {
                                     order[i] = ii + 1;
                                     excelApp.Cells[2 + i, 4] = string.Format("{0}", 0);
                                 }


                             }
                           //  excelApp.Cells[2 + i, 4] = string.Format("{0}",1);



                             }

                     }*/






                /*  // 自動調整欄寬
                  wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[5, 2]];
                  wRange.Select();
                  wRange.Columns.AutoFit();

                  //在工作簿 新增一張 統計圖表，單獨放在一個分頁裡面
                  wBook.Charts.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                  //選擇 統計圖表 的 圖表種類
                  wBook.ActiveChart.ChartType = Excel.XlChartType.xlLineMarkers;//插入折線圖
                  //設定數據範圍
                  string strRange = "A1:B3";
                  //設定 統計圖表 的 數據範圍內容
                  wBook.ActiveChart.SetSourceData(wSheet.get_Range(strRange), Excel.XlRowCol.xlColumns);
                  //將新增的統計圖表 插入到 指定位置(可以從單獨的分頁放到一個分頁裡面)
                  // wBook.ActiveChart.Location(Excel.XlChartLocation.xlLocationAsObject, wSheet.Name);

                */

                try
                {
                    //另存活頁簿
                    wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Console.WriteLine("儲存文件於 " + Environment.NewLine + pathFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("儲存檔案出錯，檔案可能正在使用" + Environment.NewLine + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("產生報表時出錯！" + Environment.NewLine + ex.Message);
            }



            //關閉活頁簿
            // wBook.Close(false, Type.Missing, Type.Missing);

            //關閉Excel
            // excelApp.Quit();
            /*
            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wSheet = null;
            wRange = null;
            excelApp = null;
            GC.Collect();*/










            #endregion

            Console.Read();
        }

        private static IConstraint AddGe(ILinearNumExpr constraint5_1, double v)
        {
            throw new NotImplementedException();
        }

        private static IConstraint AddGe(ILinearNumExpr constraint3, int cost)
        {
            throw new NotImplementedException();
        }

        public static List<List<string>> Read_Fist_Sheet(Worksheet objective_worksheet)
        {
            Console.WriteLine("Read sheet: " + objective_worksheet.Name);

            List<List<string>> result = new List<List<string>>();
            List<string> temp = new List<string>();

            int row_index = 1, column_index = 1, watch_count = 0;

            while (true)
            {
                Range current_cell = (Range)objective_worksheet.Cells[row_index, column_index];

                if (current_cell.Value != null)
                {
                    if (current_cell.Value.ToString() == "資源主檔")
                    {
                        watch_count = 0;
                        break;
                    }
                    if (watch_count > 100)
                    {
                        Console.WriteLine("Error! Failed to find the information in Sheet1");
                        break;
                    }
                    else
                    {
                        column_index++;
                        watch_count++;
                    }
                }
                else
                {
                    column_index++;
                    watch_count++;
                }
            }

            Console.WriteLine("Fist block index: (" + (row_index + 2).ToString() + "," + column_index.ToString() + ")");

            temp.Add(((Range)objective_worksheet.Cells[row_index + 3, column_index + 1]).Value.ToString()); // result[0][0] = 大線,資源數量
            temp.Add(((Range)objective_worksheet.Cells[row_index + 4, column_index + 1]).Value.ToString()); // result[0][1] = 大線,線速
            temp.Add(((Range)objective_worksheet.Cells[row_index + 3, column_index + 2]).Value.ToString()); // result[0][2] = 小線,資源數量
            temp.Add(((Range)objective_worksheet.Cells[row_index + 4, column_index + 2]).Value.ToString()); // result[0][3] = 小線,線速
            temp.Add(((Range)objective_worksheet.Cells[row_index + 3, column_index + 4]).Value.ToString()); // result[0][4] = 金油,資源數量
            temp.Add(((Range)objective_worksheet.Cells[row_index + 4, column_index + 4]).Value.ToString()); // result[0][5] = 金油,線速

            result.Add(temp);
            temp = new List<string>();
            row_index = 6;

            while (true)
            {
                Range current_cell = (Range)objective_worksheet.Cells[row_index, column_index];

                if (current_cell.Value != null)
                {

                    if (current_cell.Value.ToString() == "出發點")
                    {
                        watch_count = 0;
                        break;
                    }
                    if (watch_count > 100)
                    {
                        Console.WriteLine("Error! Failed to find the information in Sheet1");
                        break;
                    }
                    else
                    {
                        row_index++;
                        watch_count++;
                    }
                }
                else
                {
                    row_index++;
                    watch_count++;
                }

            }

            temp.Add(((Range)objective_worksheet.Cells[row_index + 1, column_index]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 1, column_index + 1]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 1, column_index + 3]).Value.ToString());
            temp.Add(((Range)objective_worksheet.Cells[row_index + 2, column_index]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 2, column_index + 1]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 2, column_index + 3]).Value.ToString());
            temp.Add(((Range)objective_worksheet.Cells[row_index + 3, column_index]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 3, column_index + 1]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 3, column_index + 3]).Value.ToString());
            temp.Add(((Range)objective_worksheet.Cells[row_index + 4, column_index]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 4, column_index + 1]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 4, column_index + 3]).Value.ToString());
            temp.Add(((Range)objective_worksheet.Cells[row_index + 5, column_index]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 5, column_index + 1]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 5, column_index + 3]).Value.ToString());
            temp.Add(((Range)objective_worksheet.Cells[row_index + 6, column_index]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 6, column_index + 1]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 6, column_index + 3]).Value.ToString());
            temp.Add(((Range)objective_worksheet.Cells[row_index + 7, column_index]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 7, column_index + 1]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 7, column_index + 3]).Value.ToString());
            temp.Add(((Range)objective_worksheet.Cells[row_index + 8, column_index]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 8, column_index + 1]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 8, column_index + 3]).Value.ToString());
            temp.Add(((Range)objective_worksheet.Cells[row_index + 9, column_index]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 9, column_index + 1]).Value.ToString() + "," + ((Range)objective_worksheet.Cells[row_index + 9, column_index + 3]).Value.ToString());

            result.Add(temp);

            Console.WriteLine("Second block index: (" + row_index.ToString() + "," + column_index.ToString() + ")\n");

            return result;
        }
        public static List<List<string>> Read_Sheet(Worksheet objective_worksheet)
        {
            List<List<string>> result = new List<List<string>>();

            Console.WriteLine("Read sheet: " + objective_worksheet.Name);

            //objective_worksheet.Columns.Count

            int row_index = 1, column_index = 1;
            int row_length = 0, column_length = 0;

            while (true)
            {
                Range current_cell = (Range)objective_worksheet.Cells[row_index, column_index];

                if (current_cell.Text != "")
                {
                    column_length++;
                    column_index++;
                }
                else
                {
                    break;
                }
            }

            Console.WriteLine("Total " + column_length.ToString() + " Columns");

            row_index = 2;

            while (true)
            {
                column_index = 1;
                string[] temp = new string[column_length];
                Range current_cell = (Range)objective_worksheet.Cells[row_index, column_index];

                if (current_cell.Text != "")
                {
                    for (int i = 1; i <= column_length; i++)
                    {
                        current_cell = (Range)objective_worksheet.Cells[row_index, i];

                        temp[i - 1] = current_cell.Text.ToString();
                    }

                    row_length++;
                }
                else
                {
                    break;
                }


                result.Add(temp.ToList());
                row_index++;
            }

            Console.WriteLine("Total " + row_length.ToString() + " Rows\n");

            return result;
        }

        public static void Initialize_List_with_Value(List<List<List<int>>> objective, List<List<int>> Job_list, int value)
        {
            for (int i = 0; i < Job_list.Count; i++)
            {
                List<List<int>> first_temp = new List<List<int>>();

                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    List<int> second_temp = new List<int>();

                    for (int k = 0; k < Machine_type; k++)
                    {
                        second_temp.Add(value);
                    }

                    first_temp.Add(second_temp);
                }

                objective.Add(first_temp);
            }
        }
        public static void Initialize_List_with_Value(List<List<int>> objective, List<List<int>> Job_list, int value)
        {
            for (int i = 0; i < Job_list.Count; i++)
            {
                List<int> temp = new List<int>();

                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    temp.Add(value);
                }

                objective.Add(temp);
            }
        }
        public static void Set_Job_Name_And_Job_Demand(List<List<string>> Data, List<string> Job_name, List<int> Job_demand)
        {
            for (int i = 0; i < Data.Count; i++)
            {
                if (!Job_name.Contains(Data[i][2]))
                {
                    Job_name.Add(Data[i][2]);
                    Job_demand.Add(int.Parse(Data[i][4]));
                }
                else
                {
                    Job_demand[Job_name.IndexOf(Data[i][2])] += int.Parse(Data[i][4]);
                }
            }
        }
        public static void Set_Job_List_And_Pitch_Required(List<List<string>> Data, List<string> Job_name, List<int> Job_demand, List<List<int>> Job_list, List<double> Pitch_required)
        {
            for (int i = 0; i < Job_name.Count; i++)
            {

                List<int> current_process = new List<int>();

                int row_index = 0;

                // find the row index in sheet2

                for (int j = 0; j < Data.Count; j++)
                {
                    if (Data[j][0] == Job_name[i])
                    {
                        row_index = j;
                        break;
                    }
                }

                // Job_list

                for (int j = 2; j < Data[row_index].Count; j++)
                {
                    string process = Data[row_index][j].Trim();

                    switch (process)
                    {
                        case "噴漆":
                            current_process.Add(0);
                            break;
                        case "貼標":
                            current_process.Add(1);
                            break;
                        case "金油線":
                            current_process.Add(2);
                            break;
                        default:
                            break;
                    }


                }

                Job_list.Add(current_process);

                // Pitch_Required

                double used_pitch_per_car = double.Parse(Data[row_index][1]);

                Pitch_required.Add(Math.Ceiling(used_pitch_per_car * Job_demand[i]));

            }
        }
        public static void Set_Transportation_time(List<string> Data, List<List<int>> Transportation_time)
        {
            //Transportation_time = new List<List<int>>();

            for (int i = 0; i < Machine_type; i++)
            {
                List<int> temp = new List<int>();

                for (int j = 0; j < Machine_type; j++)
                {
                    temp.Add(0);
                }

                Transportation_time.Add(temp);
            }

            for (int i = 0; i < Data.Count; i++)
            {
                string[] temp = Data[i].Split(',');

                int from_index = -1, to_idnex = -1;

                switch (temp[0].Trim())
                {
                    case "大線":
                        from_index = 0;
                        break;
                    case "小線":
                        from_index = 1;
                        break;
                    case "貼標":
                        from_index = 2;
                        break;
                    case "金油線":
                        from_index = 3;
                        break;
                    default:
                        break;
                }

                switch (temp[1].Trim())
                {
                    case "大線":
                        to_idnex = 0;
                        break;
                    case "小線":
                        to_idnex = 1;
                        break;
                    case "貼標":
                        to_idnex = 2;
                        break;
                    case "金油線":
                        to_idnex = 3;
                        break;
                    default:
                        break;
                }

                Transportation_time[from_index][to_idnex] = int.Parse(temp[2]);

            }
        }
        public static void Set_Process_Time(List<string> Data, List<List<string>> Data2, List<List<List<int>>> Process_time, List<List<int>> Job_list,
            List<double> Pitch_required, List<string> Job_name, List<int> Job_demand)
        {
            Initialize_List_with_Value(Process_time, Job_list, 0);

            int[] Cycle_pitch_num = new int[Machine_type] { int.Parse(Data[0]), int.Parse(Data[2]), 0, int.Parse(Data[4]) };
            int[] Cycle_rate = new int[Machine_type] { int.Parse(Data[1]), int.Parse(Data[3]), 0, int.Parse(Data[5]) };

            for (int i = 0; i < Job_list.Count; i++)
            {
                int row_index = -1;

                for (int k = 0; k < Data.Count; k++)
                {
                    if (Data2[k][0] != Job_name[i])
                    {
                        row_index = k;
                        break;
                    }
                }

                int process_time_per_unit = int.Parse(Data2[row_index][5]);

                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    if (Job_list[i][j] == 0)
                    {
                        Process_time[i][j][0] = (Cycle_pitch_num[0] + (int)Pitch_required[i]) * Cycle_rate[0];
                        Process_time[i][j][1] = (Cycle_pitch_num[1] + (int)Pitch_required[i]) * Cycle_rate[1];
                    }
                    else if (Job_list[i][j] == 1)
                    {
                        Process_time[i][j][2] = Job_demand[i] * process_time_per_unit / Crew_num;
                    }
                    else
                    {
                        Process_time[i][j][3] = (Cycle_pitch_num[3] + (int)Pitch_required[i]) * Cycle_rate[3];
                    }

                }

            }
        }
        public static void Set_Given_x_variable_solutions(List<List<List<int>>> Given_x_variable_solutions, List<List<int>> Job_list)
        {
            Initialize_List_with_Value(Given_x_variable_solutions, Job_list, -1);

            for (int i = 0; i < Job_list.Count; i++)
            {
                for (int j = 0; j < Job_list[i].Count; j++)
                {
                    if (Job_list[i][j] == 0)
                    {
                        Given_x_variable_solutions[i][j][2] = 0;
                        Given_x_variable_solutions[i][j][3] = 0;
                    }
                    else if (Job_list[i][j] == 1)
                    {
                        Given_x_variable_solutions[i][j][0] = 0;
                        Given_x_variable_solutions[i][j][1] = 0;
                        Given_x_variable_solutions[i][j][2] = 1;
                        Given_x_variable_solutions[i][j][3] = 0;
                    }
                    else
                    {
                        Given_x_variable_solutions[i][j][0] = 0;
                        Given_x_variable_solutions[i][j][1] = 0;
                        Given_x_variable_solutions[i][j][2] = 0;
                        Given_x_variable_solutions[i][j][3] = 1;

                    }
                }
            }
        }
        public static void Set_Given_e_variable_solutions(List<List<int>> Given_e_variable_solutions, List<List<int>> Job_list)
        {
            Initialize_List_with_Value(Given_e_variable_solutions, Job_list, -1);

            for (int i = 0; i < Job_list.Count; i++)
            {
                for (int j = 0; j < Job_list[i].Count - 2; j++)
                {
                    if (Job_list[i][j] == 1 || Job_list[i][j] == 2)
                    {
                        Given_e_variable_solutions[i][j] = 0;
                    }
                    else if (!(Job_list[i][j] == 0 && Job_list[i][j + 1] == 0 && Job_list[i][j + 2] == 0))
                    {
                        Given_e_variable_solutions[i][j] = 0;
                    }
                }
            }
        }
    }
}
