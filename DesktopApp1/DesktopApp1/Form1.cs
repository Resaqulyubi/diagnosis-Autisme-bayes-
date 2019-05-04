using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

// This is the code for your desktop app.
// Press Ctrl+F5 (or go to Debug > Start Without Debugging) to run your app.

namespace DesktopApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // Click on the link below to continue learning how to build a desktop app using WinForms!
            System.Diagnostics.Process.Start("http://aka.ms/dotnet-get-started-desktop");

        }

        private void button1_Click(object sender, EventArgs e)
        {

            //create a list to hold all the values
            List<string> excelData = new List<string>();
            List<List<string>> data = new List<List<string>>();
          
            //read the Excel file as byte array
            byte[] bin = File.ReadAllBytes("D:\\tes.xlsx");

            float mUserCounter = 0;
            float mcounterParah = 0;
            float mcounterSedang = 0;
            float mcounterRingan = 0;

            List<string> mUserJawaban = new List<string>();
            List<string> mKriteriaHistory = new List<string>();
            List<float> mPrior = new List<float>(); //index 0-> parah ; 1->sedang;
            List<List<float>> mPeluangOnParah = new List<List<float>>();
            List<List<float>> mPeluangOnSedang = new List<List<float>>();
            List<List<float>> mPeluangOnRingan = new List<List<float>>();

            //create a new Excel package in a memorystream
            using (MemoryStream stream = new MemoryStream(bin))
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                //loop all worksheets
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    //loop all rows
                    for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                    {
                        richTextBox1.AppendText("\n");
                        List<string> track = new List<string>();
                        //loop all columns in a row
                        for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                        {
                            
                            //add the cell data to the List
                            if (worksheet.Cells[i, j].Value != null)
                            {
                                richTextBox1.AppendText(worksheet.Cells[i, j].Value.ToString()+" | ");
                                //excelData.Add(worksheet.Cells[i, j].Value.ToString());
                                track.Add(worksheet.Cells[i, j].Value.ToString());


                            }
                        }

                        data.Add(track);
                    }
                }
            }

        
            Console.WriteLine(data);

            if (data.Count > 0)
            {
            

                for (int i = 1; i < data.Count; i++)
                {
                    for (int z = 0; z < data[i].Count; z++)
                    {

                        if(i == data.Count-1)
                        {
                            mUserJawaban.Add(data[i][z]);
                        }

                        if (z == 20)
                        {

                            mKriteriaHistory.Add(data[i][z]);
                            switch (data[i][z])
                            {
                                case "Parah":
                                    mcounterParah++;

                                    break;
                                case "Sedang":

                                    mcounterSedang++;
                                    break;
                                    case "Ringan":

                                    mcounterRingan++;
                                    break;

                                default:
                                    break;
                            }

                        }

                    }


                }

                mUserCounter = mKriteriaHistory.Count;

                float priorParah = mcounterParah / mUserCounter;
                float priorSedang = mcounterSedang / mUserCounter;
                float priorRingan = mcounterRingan / mUserCounter;

                mPrior.Add(priorParah);
                mPrior.Add(priorSedang);
                mPrior.Add(priorRingan);


                richTextBox1.AppendText("\n\n total =" + mUserCounter);
                richTextBox1.AppendText("\n\n total mcounterParah =" + mcounterParah);
                richTextBox1.AppendText("\n\n total mcounterSedang =" + mcounterSedang);
                richTextBox1.AppendText("\n\n total mcounterRingan =" + mcounterRingan);
                richTextBox1.AppendText("\n priorParah =" + priorParah);
                richTextBox1.AppendText("\n priorSedang =" + priorSedang);
                richTextBox1.AppendText("\n priorRingan =" + priorRingan);
              
                //perhitungan  peluang


                for (int i = 0; i < 2; i++)
                {
                    List<float> trackParah = new List<float>();
                    List<float> trackSedang = new List<float>();
                    List<float> trackRingan = new List<float>();
                    for (int z = 0; z < 20; z++)
                    {
                        trackParah.Add(0);
                        trackSedang.Add(0);
                        trackRingan.Add(0);
                    }

                    mPeluangOnParah.Add(trackParah);
                    mPeluangOnSedang.Add(trackSedang);
                    mPeluangOnRingan.Add(trackRingan);
                }

               

                for (int i = 1; i < data.Count-1; i++)
                {
                    for (int z = 0; z < data[i].Count-1; z++)
                    {
                         if(data[i][z].Equals(mUserJawaban[z]))
                        {

                            switch (mKriteriaHistory[i-1])
                            {
                                case "Parah":
                                    mPeluangOnParah[0][z] = mPeluangOnParah[0][z] +1;
                                
                                    break;

                                case "Sedang":
                                    mPeluangOnSedang[0][z] = mPeluangOnSedang[0][z]+1;

                                    break;
                                case "Ringan":
                                    mPeluangOnRingan[0][z] = mPeluangOnRingan[0][z]+1;

                                    break;
                            }

                            
                        }
                    }
                }

                richTextBox1.AppendText("\n ========================================= \n");

                for (int i = 0; i < mPeluangOnParah[0].Count; i++)
                {
                    richTextBox1.AppendText("\n mJumlahPeluangOnParah " +" F"+ (i+1)+" ="+ mPeluangOnParah[0][i]);
                    richTextBox1.AppendText("\n mJumlahPeluangOnSedang " + " F"+ (i+1)+" ="+ mPeluangOnSedang[0][i]);
                    richTextBox1.AppendText("\n mJumlahPeluangOnRingan " + " F"+ (i+1)+" ="+ mPeluangOnRingan[0][i]);

                    mPeluangOnParah[1][i] = mPeluangOnParah[0][i] / mPrior[0];
                    mPeluangOnSedang[1][i] = mPeluangOnSedang[0][i] / mPrior[1];
                    mPeluangOnRingan[1][i] = mPeluangOnRingan[0][i] / mPrior[2];
              

                }



                float totalPeluangParah = 0;
                float totalPeluangSedang = 0;
                float totalPeluangRingan = 0;

                richTextBox1.AppendText("\n ========================================= \n");

                for (int i = 0; i < mPeluangOnParah[1]. Count; i++)
                {
                    richTextBox1.AppendText("\n mPeluangOnParah " + " F" + (i + 1) + " =" + mPeluangOnParah[1][i]);
                    richTextBox1.AppendText("\n mPeluangOnSedang " + " F" + (i + 1) + " =" + mPeluangOnSedang[1][i]);
                    richTextBox1.AppendText("\n mPeluangOnRingan " + " F" + (i + 1) + " =" + mPeluangOnRingan[1][i]);


                    if (i == 0)
                    {
                        totalPeluangParah = mPeluangOnParah[1][i];
                        totalPeluangSedang = mPeluangOnSedang[1][i];
                        totalPeluangRingan = mPeluangOnRingan[1][i];
                    }
                    else
                    {
                        totalPeluangParah = totalPeluangParah * mPeluangOnParah[1][i];
                        totalPeluangSedang = totalPeluangSedang * mPeluangOnSedang[1][i];
                        totalPeluangRingan = totalPeluangRingan * mPeluangOnRingan[1][i];
                    }



                }

                richTextBox1.AppendText("\n ========================================= \n");


                richTextBox1.AppendText("\n totalPeluangParah =" + totalPeluangParah);
                richTextBox1.AppendText("\n totalPeluangSedang  =" + totalPeluangSedang);
                richTextBox1.AppendText("\n totalPeluangRingan  =" + totalPeluangRingan);

                float mResultParah = totalPeluangParah *mPrior[0];
                float mResultSedang = totalPeluangSedang * mPrior[1];
                float mResultRingan = totalPeluangRingan * mPrior[2];

                richTextBox1.AppendText("\n mResultParah =" + mResultParah);
                richTextBox1.AppendText("\n mResultSedang  =" + mResultSedang);
                richTextBox1.AppendText("\n mResultRingan  =" + mResultRingan);


                String resultKriteria = "";

                float a=   Math.Max(mResultParah, mResultSedang);
                float b=   Math.Max(a, mResultRingan);



                if (b == mResultParah)
                {
                    resultKriteria = "Parah";
                }
                else if (b == mResultSedang)
                {
                    resultKriteria = "Sedang";
                }
                else
                {
                    resultKriteria = "Ringan";
                }

           
                richTextBox1.AppendText("\n ========================================= \n");

                richTextBox1.AppendText("\n hasil kriteria yang diperoleh   =" + resultKriteria);


            }



            Console.WriteLine(value: "log");
                
        }
    }
} 
