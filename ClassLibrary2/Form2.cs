using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.Data.OleDb;
using DataTable = System.Data.DataTable;
using System.IO;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.EditorInput;

namespace ClassLibrary2
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            dataGridView1.ColumnCount = 4;
            dataGridView1.Columns[0].Name = "x";
            dataGridView1.Columns[1].Name = "y";
            dataGridView1.Columns[2].Name = "z";
            dataGridView1.Columns[3].Name = "grade";
   
        }
        double x_min = 1000000, x_max = -1000000;
        double y_min = 1000000, y_max = -1000000;
        double z_min = 1000000, z_max = -1000000;
        int i = 0;    

        public void add(string[] row)
        {
            dataGridView1.Rows.Add(row);
            if(x_min > Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))
            {
                x_min = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
            }
            if (x_max < Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))
            {
                x_max = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
            }
            if (y_min > Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value))
            {
                y_min = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
            }
            if (y_max < Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value))
            {
                y_max = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
            }
            if (z_min > Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value))
            {
                z_min = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
            }
            if (z_max < Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value))
            {
                z_max = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
            }
            i = i + 1;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.interpolation();
           // this.drawCube();
            this.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        public void interpolation()
        {
            //For color of cubes according to thier grade values....
            Autodesk.AutoCAD.Colors.Color[] acColors = new Autodesk.AutoCAD.Colors.Color[9];
            acColors[0] = Autodesk.AutoCAD.Colors.Color.FromRgb(201, 248, 227);
            acColors[1] = Autodesk.AutoCAD.Colors.Color.FromRgb(245, 196, 224);
            acColors[2] = Autodesk.AutoCAD.Colors.Color.FromRgb(243, 11, 220);
            acColors[3] = Autodesk.AutoCAD.Colors.Color.FromRgb(11, 193, 213);
            acColors[4] = Autodesk.AutoCAD.Colors.Color.FromRgb(11, 243, 173);
            acColors[5] = Autodesk.AutoCAD.Colors.Color.FromRgb(139, 243, 11);
            acColors[6] = Autodesk.AutoCAD.Colors.Color.FromRgb(243, 208, 11);
            acColors[7] = Autodesk.AutoCAD.Colors.Color.FromRgb(243, 123, 11);
            acColors[8] = Autodesk.AutoCAD.Colors.Color.FromRgb(243, 11, 11);

            for (double z = z_min; z <= z_max; z = z + 2)
            {
                for (double y = y_min; y <= y_max; y = y + 2)
                {
                    for (double x = x_min; x <= x_max; x = x + 2)
                    {
                        List<double[]> myarray = new List<double[]>();
                        calculate_relev_distance(x, y, z,ref myarray);
                        double grade =  calculate_grade(ref myarray);
                        if(grade > 0)
                        {
                            drawPoint(x,y,z,grade,ref acColors);
                        }
                        myarray.Clear();
                    }
                }
            }
        }

        public void drawPoint(double x, double y, double z, double grade, ref Autodesk.AutoCAD.Colors.Color[] acColors)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.DatabaseServices.Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            // Start a transaction
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    // Open the Block table record for read
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                 OpenMode.ForRead) as BlockTable;

                    // Open the Block table record Model space for write
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                    OpenMode.ForWrite) as BlockTableRecord;

                    /* Create a single-line text object
                    using (DBText acText = new DBText())
                    {
                        acText.Position = new Point3d(x, y, z);
                        acText.Height = 0.5;
                        acText.TextString = grade.ToString("0.###");
                        acBlkTblRec.AppendEntity(acText);
                        acTrans.AddNewlyCreatedDBObject(acText, true);
                    } */                   

                    // Create a 3D solid box
                    Solid3d acSol3D = new Solid3d();
                    acSol3D.SetDatabaseDefaults();
                    acSol3D.CreateBox(2, 2, 2);
                    // Position the center of the 3D solid at (5,5,0) 
                    acSol3D.TransformBy(Matrix3d.Displacement(new Point3d(x, y, z) - Point3d.Origin));

                    //Color
                    Autodesk.AutoCAD.Colors.Color curcolor = new Autodesk.AutoCAD.Colors.Color();

                    if (grade.ToString("0.#") == "0.9")
                        { curcolor = acColors[8]; }
                    else if (grade.ToString("0.#") == "0.8")
                        { curcolor = acColors[7]; }
                    else if (grade.ToString("0.#") == "0.7")
                        { curcolor = acColors[6]; }
                    else if (grade.ToString("0.#") == "0.6")
                        { curcolor = acColors[5]; }
                    else if (grade.ToString("0.#") == "0.5")
                        { curcolor = acColors[4]; }
                    else if (grade.ToString("0.#") == "0.4")
                        { curcolor = acColors[3]; }
                    else if (grade.ToString("0.#") == "0.3")
                        { curcolor = acColors[2]; }
                    else if (grade.ToString("0.#") == "0.2")
                        { curcolor = acColors[1]; }
                    else if (grade.ToString("0.#") == "0.1")
                        { curcolor = acColors[0]; }
                    
                    acSol3D.Color = curcolor;

                    //Block thickness
                    acSol3D.LineWeight = LineWeight.LineWeight020;
                    acCurDb.LineWeightDisplay = true;

                    // Add the new object to the block table record and the transaction
                    acBlkTblRec.AppendEntity(acSol3D);
                    acTrans.AddNewlyCreatedDBObject(acSol3D, true);
                    acTrans.Commit();
                }
            }
        }

        public double calculate_grade(ref List<double[]> myarray)
        {
            double sum_up = 0;
            double sum_down = 0;
            for(int i = 0; i < myarray.Count(); i++)
            {
                if (myarray[i][0] == 0)
                    return myarray[i][1];
                sum_up = sum_up + myarray[i][1] / (myarray[i][0]* myarray[i][0]);
                sum_down = sum_down + (1 / (myarray[i][0] * myarray[i][0]));
            }
            if (sum_down == 0)
                return 0;

            double grade = sum_up / sum_down;
            return grade;
        }

        public void calculate_relev_distance(double x,double y,double z,ref List<double[]> myarray)
        {
            for (i = 0; i < dataGridView1.Rows.Count ; i++)
            {
                double x1 = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                double y1 = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                double z1 = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                double g = Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);

                double d = Math.Sqrt((x1 - x) * (x1 - x) + (y1 - y) * (y1 - y) + (z1 - z) * (z1 - z));
                if(d < 10)   //Setting inference zone
                {
                    double[] dis = new double[] { d, g };
                    myarray.Add(dis);
                }
            }
        }
    }
}
