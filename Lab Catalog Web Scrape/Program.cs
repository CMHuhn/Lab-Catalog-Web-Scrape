using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace LabCatalogWebScrape
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\Users\Charles\Documents\Visual Studio 2015\Projects\Lab Catalog Web Scrape\testinfo.txt";
            string mainurl = "http://www.sanfordhealth.org/bismarck-nd/labtestcatalog/listall.asp";
            string mainresult = null;
            WebResponse mainresponse = null;
            StreamReader mainreader = null;
            var csv = new StringBuilder();
            int i = 0;

            try
            {
                HttpWebRequest mainrequest = (HttpWebRequest)WebRequest.Create(mainurl);
                mainrequest.Method = "GET";
                mainresponse = mainrequest.GetResponse();
                mainreader = new StreamReader(mainresponse.GetResponseStream(), Encoding.UTF8);
                mainresult = mainreader.ReadToEnd();
            }
            catch (Exception ex)
            {
                // handle error
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (mainreader != null)
                    mainreader.Close();
                if (mainresponse != null)
                    mainresponse.Close();
            }


            Match mainm;
            string mainpat = "(?<1>testDetail\\.asp\\?CID=\\d*)";

            mainm = Regex.Match(mainresult, mainpat);

            while (mainm.Success)
            {
                string url = String.Concat("http://www.sanfordhealth.org/bismarck-nd/labtestcatalog/", mainm.Groups[1].Value);
                string result = null;
                WebResponse response = null;
                StreamReader reader = null;

                try
                {
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                    request.Method = "GET";
                    response = request.GetResponse();
                    reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                    result = reader.ReadToEnd();
                }
                catch (Exception ex)
                {
                    // handle error
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (reader != null)
                        reader.Close();
                    if (response != null)
                        response.Close();
                }

                Match m;
                string tnpat = "Test name:\\S+\\s+<td width=\"81%\"\\s*\\S*><h3>\\s*(?<1>.*)</h3>";
                string tcpat = "Test code:\\S+\\s+<td\\s*\\S*>\\s*(?<2>.*?)\\s*</td>";
                string ocpat = "Order code:\\S+\\s+<td\\s*\\S*>\\s*(?<3>.*?)\\s*</td>";
                string cptpat = "CPT:\\S+\\s+<td\\s*\\S*>\\s*(?<4>.*?)\\s*</td>";
                string stpat = "Specimen type:\\S+\\s+<td\\s*\\S*>\\s*(?<5>.*?)\\s*</td>";
                string copat = "Container:\\S+\\s+<td\\s*\\S*>\\s*(?<6>.*?)\\s*</td>";
                string svpat = "Specimen volume:\\S+\\s+<td\\s*\\S*>\\s*(?<7>.*?)\\s*</td>";
                string inpat = "Instructions:\\S+\\s+<td\\s*\\S*>\\s*(?<8>[\\s\\S]*?)</td>";
                string nopat = "Note:\\S+\\s+<td\\s*\\S*>\\s*(?<9>.*?)\\s*</td>";
                string trpat = "Transport:\\S+\\s+<td\\s*\\S*>\\s*(?<10>.*?)\\s*</td>";
                string smpat = "Specimen minimum <br />volume:\\S+\\s+<td\\s*\\S*>\\s*(?<11>.*?)\\s*</td>";
                string pepat = "Performed:\\S+\\s+<td\\s*\\S*>\\s*(?<12>.*?)\\s*</td>";
                string rvpat = "Reference value:\\S+\\s+<td\\s*\\S*>\\s*(?<13>.*?)\\s*</td>";
                string mepat = "Method:\\S+\\s+<td\\s*\\S*>\\s*(?<14>.*?)\\s*</td>";

                string testname, testcode, ordercode, cpt, specimentype, container, specimenvolume, instructions, note, transport, specimenminimum, performed, referencevalue, method;

                m = Regex.Match(result, tnpat);
                if (m.Success) testname = m.Groups[1].Value;
                else testname = "";

                m = Regex.Match(result, tcpat);
                if (m.Success) testcode = m.Groups[2].Value;
                else testcode = "";

                m = Regex.Match(result, ocpat);
                if (m.Success) ordercode = m.Groups[3].Value;
                else ordercode = "";

                m = Regex.Match(result, cptpat);
                if (m.Success) cpt = m.Groups[4].Value;
                else cpt = "";

                m = Regex.Match(result, stpat);
                if (m.Success) specimentype = m.Groups[5].Value;
                else specimentype = "";

                m = Regex.Match(result, copat);
                if (m.Success) container = m.Groups[6].Value;
                else container = "";

                m = Regex.Match(result, svpat);
                if (m.Success) specimenvolume = m.Groups[7].Value;
                else specimenvolume = "";

                m = Regex.Match(result, inpat);
                if (m.Success) instructions = m.Groups[8].Value;
                else instructions = "";
                //instructions = instructions.Replace("\r\n", "").Replace("\n", "").Replace("\r", "");

                m = Regex.Match(result, nopat);
                if (m.Success) note = m.Groups[9].Value;
                else note = "";

                m = Regex.Match(result, trpat);
                if (m.Success) transport = m.Groups[10].Value;
                else transport = "";

                m = Regex.Match(result, smpat);
                if (m.Success) specimenminimum = m.Groups[11].Value;
                else specimenminimum = "";

                m = Regex.Match(result, pepat);
                if (m.Success) performed = m.Groups[12].Value;
                else performed = "";

                m = Regex.Match(result, rvpat);
                if (m.Success) referencevalue = m.Groups[13].Value;
                else referencevalue = "";

                m = Regex.Match(result, mepat);
                if (m.Success) method = m.Groups[14].Value;
                else method = "";
                
                /* 1: Test Name
                 * 2: AKA
                 * 3: Order Code
                 * 4: Container Type
                 * 5: Specimen type/requirement
                 * 6: Specimen volume
                 * 7: Minimum volume
                 * 8: Stability
                 * 9: Interfering substances
                 * 10: Transport Temp
                 * 11: Instructions
                 * 12: Performed Test Frequency
                 * 13: Methodology
                 * 14: Performing lab
                 * 15: CPT
                 */

                string wspattern = "\\s+";
                string space = " ";
                Regex wsreplace = new Regex(wspattern);
                instructions = wsreplace.Replace(instructions, space);

                string bracketpattern = "<.*?>";
                Regex bracketreplace = new Regex(bracketpattern);
                instructions = bracketreplace.Replace(instructions, string.Empty);

                var newLine = string.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11};{12};{13}", testname, testcode, ordercode, cpt, specimentype, container, specimenvolume, instructions, note, transport, specimenminimum, performed, referencevalue, method);
                csv.AppendLine(newLine);
                mainm = mainm.NextMatch();
                System.Console.WriteLine(i++);

                //if (i > 25) break;
            }

            if (!File.Exists(path))
            {
                File.WriteAllText(path, csv.ToString());
            }
        }
    }
}
