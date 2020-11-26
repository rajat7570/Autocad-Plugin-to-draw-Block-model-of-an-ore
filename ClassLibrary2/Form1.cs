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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button3_Click_1(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string path;
                path = openFileDialog1.FileName;
                string ext = Path.GetExtension(path);
                if (ext == ".csv")
                {
                    DataTable my_data = new DataTable();
                    string[] raw_text = File.ReadAllLines(path);
                    string[] data_col = null;
                    int x = 0;

                    foreach (string test_line in raw_text)
                    {
                        //MessageBox.Show(test_line);
                        data_col = test_line.Split(',');
                        if (x == 0)
                        {
                            //header
                            for (int i1 = 0; i1 < data_col.Count(); i1++)
                            {
                                my_data.Columns.Add(data_col[i1]);
                                //dataGridView1.Columns.Add(data_col[i1]);
                            }
                            x++;
                        }
                        else
                        {
                            //data
                            my_data.Rows.Add(data_col);
                        }

                    }
                    dataGridView2.DataSource = my_data;
                    int a, i;
                    a = dataGridView2.Columns.Count;
                    for (i = 0; i < a; i++)
                    {
                        dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;

                    }
                }
            }
        }
        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            //this.draw();
            this.draw_line();
            this.draw_scale();
            this.Close();
        }

        
        void draw_scale()
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

            int x1 = 30;
            int y1 = 5;
            int z1 = 0;

            int x2 = x1 + 2;
            int y2 = 5;
            int z2 = 0;
            for (int i = 0; i < 9; i++)
            {
                int d = 3;
                int x3 = x1;
                int y3 = y1 - d;
                int z3 = 0;

                int x4 = x2;
                int y4 = y3;
                int z4 = 0;

                Point3d p0 = new Point3d(x1, y1, z1);
                Point3d p1 = new Point3d(x2, y2, z2);
                Point3d p2 = new Point3d(x3, y3, z3);
                Point3d p3 = new Point3d(x4, y4, z4);

                Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.DatabaseServices.Database acCurDb = acDoc.Database;
                Editor ed = acDoc.Editor;
                // Start a transaction
                using (DocumentLock docLock = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                    {

                        BlockTable acBlkTbl;
                        acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                        OpenMode.ForRead) as BlockTable;
                        // Open the Block table record Model space for write
                        BlockTableRecord acBlkTblRec;
                        acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                        OpenMode.ForWrite) as BlockTableRecord;

                        using (Solid ac2DSolidSqr1 = new Solid(p0, p1, p2, p3))
                        {
                            //ac2DSolidSqr1.Color = acColors[i];
                            Autodesk.AutoCAD.Colors.Color curcolor = new Autodesk.AutoCAD.Colors.Color();                            
                            curcolor = acColors[8-i];                             
                            ac2DSolidSqr1.Color = curcolor;

                            // Add the new object to the block table record and the transaction
                            acBlkTblRec.AppendEntity(ac2DSolidSqr1);
                            acTrans.AddNewlyCreatedDBObject(ac2DSolidSqr1, true);
                        }

                        // Create a single-line text object
                        using (DBText acText = new DBText())
                        {
                            acText.Position = new Point3d(x4 + 1, (y1+y3)/2, 0);
                            acText.Height = 0.7;
                            String grade = "0." + (10 - i-1) + " - " + "0." +(10-i);
                            acText.TextString = grade;

                            acBlkTblRec.AppendEntity(acText);
                            acTrans.AddNewlyCreatedDBObject(acText, true);
                        }
                        y1 = y3;
                        y2 = y1;
                        x1 = x3;
                        x2 = x4;
                        z1 = z3;
                        z2 = z4;

                        if (i == 5)
                        {
                            // Create a single-line text object
                            using (DBText acText = new DBText())
                            {
                                acText.Position = new Point3d(30, 6, 0);
                                acText.Height = 1;
                                acText.TextString = "Color Scale";
                                acBlkTblRec.AppendEntity(acText);
                                acTrans.AddNewlyCreatedDBObject(acText, true);
                            }
                            using (DBText acText = new DBText())
                            {
                                acText.Position = new Point3d(38, 6, 0);
                                acText.Height = 0.8;
                                acText.TextString = "(Color representation of grade)";
                                acBlkTblRec.AppendEntity(acText);
                                acTrans.AddNewlyCreatedDBObject(acText, true);
                            }
                        }                        

                        // Changing Visual style
                        DBDictionary dict = (DBDictionary)acTrans.GetObject(acCurDb.VisualStyleDictionaryId, OpenMode.ForRead);
                        ViewportTable vt = (ViewportTable)acTrans.GetObject(acCurDb.ViewportTableId, OpenMode.ForRead);
                        ViewportTableRecord vtr = (ViewportTableRecord)acTrans.GetObject(vt["*Active"], OpenMode.ForWrite);
                        vtr.VisualStyleId = dict.GetAt("Shaded with edges");

                        acTrans.Commit();
                    }
                    ed.UpdateTiledViewportsFromDatabase();
                }
            }

        }

        void draw_line()
        {
            int i;
            double xi = 0;
            double yi = 0;
            double x1 = 0;
            double y1 = 0;
            double z1 = 0;
            int L = 0;

            string prev = Convert.ToString(dataGridView2.Rows[0].Cells[1].Value);
            var BHL_data = new ClassLibrary2.Form2();
            BHL_data.Show();
            for (i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                int count = 0;
                int len = Convert.ToInt32(dataGridView2.Rows[i].Cells[1].Value);
                double angle_alpha = Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value);     // by z- axis
                double angle_beta = Convert.ToDouble(dataGridView2.Rows[i].Cells[4].Value);      // by x- axis
                Point3d p0 = new Point3d(x1, y1, z1);

                L =  len;
                double x2 = -L*Math.Sin(angle_alpha)*Math.Cos(angle_beta);
                double y2 = -L * Math.Sin(angle_alpha) * Math.Sin(angle_beta); ;
                double z2 = -L*Math.Cos(angle_alpha);
                Point3d p1 = new Point3d(x2+x1, y2+y1, z2+z1);

                if (Convert.ToString(dataGridView2.Rows[i].Cells[2].Value) == "ore")
                {   
                    double grade = Convert.ToDouble(dataGridView2.Rows[i].Cells[5].Value);          //grade
                    string[] row = new string[] {x1.ToString("0.###"),y1.ToString("0.###"), z1.ToString("0.###"), grade.ToString("0.###") } ;
                    
                    BHL_data.add(row);
                }

                Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.DatabaseServices.Database acCurDb = acDoc.Database;
                // Start a transaction
                using (DocumentLock docLock = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                    {
                        BlockTable acBlkTbl;
                        acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                        OpenMode.ForRead) as BlockTable;
                        // Open the Block table record Model space for write
                        BlockTableRecord acBlkTblRec;
                        acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                        OpenMode.ForWrite) as BlockTableRecord;

                        Line ln = new Line(p0, p1);
                        ln.ColorIndex = 7;
                        // set the lineweight for it
                        ln.LineWeight = LineWeight.LineWeight020;
                        acCurDb.LineWeightDisplay = true;

                        acBlkTblRec.AppendEntity(ln);
                        acTrans.AddNewlyCreatedDBObject(ln, true);

                        // Create a point at (4, 3, 0) in Model space
                        using (DBPoint acPoint = new DBPoint(p1))
                        {
                            // Add the new object to the block table record and the transaction
                            acBlkTblRec.AppendEntity(acPoint);
                            acTrans.AddNewlyCreatedDBObject(acPoint, true);
                        }
                        // Set the style for all point objects in the drawing
                        acCurDb.Pdmode = 34;
                        acCurDb.Pdsize = 1;

                        x1 = x1 + x2;
                        y1 = y1 + y2;
                        z1 = z1 + z2;

                        if (Convert.ToString(dataGridView2.Rows[i].Cells[0].Value) != Convert.ToString(dataGridView2.Rows[i + 1].Cells[0].Value))
                        {

                          //  x1 = xi + 5;
                          //  xi = x1;
                            //  y1 = 0;
                            //   z1 = 0;
                            x1 = Convert.ToDouble(dataGridView2.Rows[i+1].Cells[6].Value);
                            y1 = Convert.ToDouble(dataGridView2.Rows[i+1].Cells[7].Value);
                            z1 = Convert.ToDouble(dataGridView2.Rows[i+1].Cells[8].Value);
                            count = count + 1;
                            L = 0;
                        }

                        if (count > 0)
                        {
                            // Create a single-line text object
                            using (DBText acText = new DBText())
                            {
                                acText.Position = new Point3d(x: xi, y: yi, z: 1);
                                xi = x1;
                                yi = y1;
                                acText.Height = 1;
                                acText.TextString = Convert.ToString(dataGridView2.Rows[i].Cells[0].Value);

                                acBlkTblRec.AppendEntity(acText);
                                acTrans.AddNewlyCreatedDBObject(acText, true);
                            }
                            // Save the new object to the database
                            // acTrans.Commit();
                        }

                        acTrans.Commit();
                    }
                }
            }
        }
        
    }
}
