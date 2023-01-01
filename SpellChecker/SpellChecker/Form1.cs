
using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace SpellChecker
{
    public partial class MainForm : Form
    {
        #region MEMBER : INIT || FORM
        private string m_versionName = "맞춤법 검사기 (v.1.2_2022.10.05)";
        private string m_exceptWord = "<SYSTEM_16px>,<SYSTEM_14px>,<Narrative_16px>,<Narrative_14px>,<Narrative_12px>,</>,%user%,%USER%";
        private string[] m_exceptWord_array;

        private string m_filePath = "";
        private string m_resultPath = "";
        private string m_resultFinalPath = "";
        private string m_sheetName = "Sheet1"; //불러올 시트 이름
        private int m_default_startColumn = 2; //기준 열에 대한 초기화
        private int m_default_startRow = 2; //시작하는 기준 행
        private int m_defualt_endRow = -1; //-1 이면 끝까지 조사함

        private string m_busan_url = "http://speller.cs.pusan.ac.kr/"; //크롤링 대상 홈페이지

        private DateTime m_startTime; //시작 - 종료 시간 계산을 위한 멤버
        Random m_random;
        private int m_nowGap = 500; //Random 시간차를 위한 멤버
        #endregion

        #region MEMBER : DATA
        private Microsoft.Office.Interop.Excel.Application m_application;
        private Workbook m_workbook;
        private IWebDriver m_driver;

        List<DialogClass> m_dialogClass_list = new List<DialogClass>(); //대상 텍스트 담을 List
        #endregion

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            InitFormName();
            InitInputBox();
        }

        private void input_ExceptText_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void DoSelenium()
        {
            m_driver.Navigate().GoToUrl(m_busan_url);
            CheckValid(m_driver);
            CheckAndWaitForException(m_driver, CheckType.CantFindInputForm);

            foreach (DialogClass dc in m_dialogClass_list)
            {
                if (dc.m_description == "" || dc.m_description == null)
                {
                    Console.WriteLine($"DialogClass({dc.m_ID}) 는 비어있습니다.");
                    continue;
                }

                Thread.Sleep(GetRandomTimeByValue(m_nowGap * 4, m_nowGap * 2));
                CheckValid(m_driver);
                m_driver.FindElement(By.Id("tdBody")).FindElement(By.Name("inputForm")).FindElement(By.Name("text1")).SendKeys(dc.m_description);
                Thread.Sleep(GetRandomTimeByValue(m_nowGap / 10, m_nowGap / 10));
                m_driver.FindElement(By.Id("tdBody")).FindElement(By.Id("btnCheck")).Click();
                Thread.Sleep(GetRandomTimeByValue(m_nowGap * 4, m_nowGap * 2));
                CheckValid(m_driver);

                try
                {
                    List<IWebElement> element_resultList = m_driver.FindElement(By.Id("divCorrectionTableBox1st")).FindElements(By.ClassName("tableErrCorrect")).ToList<IWebElement>();
                    foreach (IWebElement reusltElement in element_resultList)
                    {
                        Thread.Sleep(GetRandomTimeByValue(m_nowGap / 5, m_nowGap / 2));
                        dc.m_result_list.Add(
                        new ResultClass(
                            targetString: reusltElement.FindElement(By.ClassName("tdErrWord")).Text,
                            checkString: reusltElement.FindElement(By.ClassName("replaceWord")).Text,
                            help: reusltElement.FindElement(By.ClassName("tdETNor")).Text)
                        );
                    }
                }
                catch (NoSuchElementException noSuch) { Console.WriteLine("[System : Check] 맞춤법 검사결과 이슈 없음."); }
                catch { }

                dc.m_isParsed = true;
                CheckValid(m_driver);
                m_driver.FindElement(By.Id("btnRenew2")).Click();
            }
        }
        private void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }
        private void SetExceptWordArray()
        {
            string tempString = m_exceptWord;
            tempString.Replace(" ", "");
            m_exceptWord_array = tempString.Split(',');
        }
        private string GetStringWithoutExceptWord(Range target)
        {
            //Console.WriteLine("[Test] "+target.Value2.ToString()); //for test
            if (target == null)
            {
                Console.WriteLine("[Test] target 이 null 입니다.");
            }
            if (target.Value2.ToString() == "" || target.Value2.ToString() == null)
            {
                return "";
            }
            if (m_exceptWord_array == null)
            {
                MessageBox.Show($"[System : error!!!] 제외 단어가 모두 비어있습니다.");
            }

            foreach (string exceptWord in m_exceptWord_array)
            {
                //returnString.Replace(exceptWord, string.Empty); //삭제 진행
                target.Replace(exceptWord, "");
            }
            //for test
            //MessageBox.Show($"[Test] 반환 스트링 : {returnString}");
            return target.Value2.ToString();
        }
        private void CheckAndWaitForException(IWebDriver m_driver, CheckType checkType)
        {
            switch (checkType)
            {
                case CheckType.CantFindInputForm:
                    {
                        if (m_driver.FindElement(By.Id("tdBody")).FindElement(By.Name("inputForm")).FindElement(By.Name("text1")) == null)
                        {
                            while (m_driver.FindElement(By.Id("tdBody")).FindElement(By.Name("inputForm")).FindElement(By.Name("text1")) == null)
                            {
                                Thread.Sleep(GetRandomTimeByValue(m_nowGap * 4, m_nowGap * 2));
                                if (m_driver.FindElement(By.Id("reload-button")) != null) m_driver.FindElement(By.Id("reload-button")).Click();
                                else m_driver.Navigate().Refresh();

                                Console.WriteLine("[System : Valid] Can't FindInputForm"); //for test
                            }
                        }
                        break;
                    }
                case CheckType.NetError:
                    {
                        List<IWebElement> element_findNetError = m_driver.FindElements(By.ClassName("neterror")).ToList<IWebElement>();
                        if (element_findNetError == null || element_findNetError.Count != 0)
                        {
                            while (element_findNetError == null || element_findNetError.Count != 0)
                            {
                                m_driver.Navigate().Refresh();
                                element_findNetError = m_driver.FindElements(By.ClassName("neterror")).ToList<IWebElement>();
                                Thread.Sleep(GetRandomTimeByValue(m_nowGap * 4, m_nowGap * 2));

                                Console.WriteLine($"[System : Valid] NetError / count = {element_findNetError.Count}"); //for test
                            }
                        }
                        break;
                    }
                case CheckType.ServiceUnavailable:
                    {
                        if (m_driver.Title == "503 Service Unavailable")
                        {
                            while (m_driver.Title == "503 Service Unavailable")
                            {
                                Thread.Sleep(GetRandomTimeByValue(m_nowGap * 6, m_nowGap * 2));
                                m_driver.Navigate().Refresh();

                                Console.WriteLine("[System : Valid] 503 Service Unavailable");

                                Thread.Sleep(GetRandomTimeByValue(m_nowGap * 2, m_nowGap));
                            }
                        }
                        break;
                    }
                case CheckType.BadGateway:
                    {
                        if (m_driver.Title == "502 Bad Gateway")
                        {
                            while (m_driver.Title == "502 Bad Gateway")
                            {
                                Thread.Sleep(GetRandomTimeByValue(m_nowGap * 6, m_nowGap * 2));
                                m_driver.Navigate().Refresh();

                                Console.WriteLine("[System : Valid] 502 Bad Gateway"); //for test
                                Thread.Sleep(GetRandomTimeByValue(m_nowGap * 2, m_nowGap));
                            }
                        }
                        break;
                    }
                default:
                    {
                        Thread.Sleep(GetRandomTimeByValue(m_nowGap * 10, m_nowGap * 5));
                        m_driver.Navigate().Refresh();

                        Console.WriteLine("[System : Valid] Default 10초 후 새로고침 한 번"); //for test
                        Thread.Sleep(GetRandomTimeByValue(m_nowGap * 2, m_nowGap));

                        break;
                    }
            }
        }
        private void CheckValid(IWebDriver m_driver)
        {
            CheckAndWaitForException(m_driver, CheckType.NetError);
            CheckAndWaitForException(m_driver, CheckType.BadGateway);
            CheckAndWaitForException(m_driver, CheckType.ServiceUnavailable);
        }
        private void RemoveObject(object obj)
        {
            if (obj == null) return;
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("[System : error!!!] 메모리 할당을 해제하는 중 문제가 발생하였습니다." + ex.ToString(), "경고!");
            }
            finally
            {
                GC.Collect();
            }
        }
        private int GetRandomTimeByValue(int value, int gap = 500)
        {
            return m_random.Next(value, value + gap);
        }
        private string GetNowTime()
        {
            string nowTime = $"{DateTime.Now.Month.ToString()}." +
                $"{DateTime.Now.Day.ToString()}." +
                $"{DateTime.Now.Hour.ToString()}." +
                $"{DateTime.Now.Minute.ToString()}." +
                $"{DateTime.Now.Second.ToString()}";

            return nowTime;
        }
        private void SetResultSheet()
        {
            Worksheet worksheet = m_workbook.Worksheets.Add(Type.Missing, m_workbook.Worksheets[1]);
            worksheet.Name = "Result";

            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Description";
            worksheet.Cells[1, 3].Value = "NeedCheck";

            ResizeSheet(worksheet);

            for (int i = 0; i < m_dialogClass_list.Count; i++)
            {
                if (m_dialogClass_list[i].m_isParsed == false) { continue; } // parsed 되지 않았다면 continue

                worksheet.Cells[i + 2, 1].Value = m_dialogClass_list[i].m_ID;
                worksheet.Cells[i + 2, 2].Value = m_dialogClass_list[i].m_description;
                worksheet.Cells[i + 2, 3].Value = m_dialogClass_list[i].m_needCheck.ToString();

                for (int j = 0; j < m_dialogClass_list[i].m_result_list.Count; j++)
                {
                    ResultClass rc = m_dialogClass_list[i].m_result_list[j];
                    string resultValue = $"\n*입력 내용 : {rc.m_targetString}\n\n*대치어 : {rc.m_checkString}\n\n*도움말 : {rc.m_help}\n";
                    worksheet.Cells[i + 2, 4 + j].Value = resultValue;
                }
            }
            m_resultPath = $"_result_{GetNowTime()}.xl";
            m_resultFinalPath = m_filePath.Replace(".xl", m_resultPath);
            m_workbook.SaveAs(m_resultFinalPath);

            RemoveObject(worksheet);
        }
        private void ResizeSheet(Worksheet worksheet)
        {
            Range widthController = worksheet.get_Range("B:B", System.Type.Missing);
            widthController.EntireColumn.ColumnWidth = 40;
            widthController = worksheet.get_Range("D:H", System.Type.Missing);
            widthController.EntireColumn.ColumnWidth = 50;
            widthController.WrapText = true;
            widthController.Font.Size = 9;

            Range indexController = worksheet.get_Range("A1:Z1", System.Type.Missing);
            indexController.Font.Size = 10;
            indexController.Font.Bold = true;
        }
        internal void EndSelenium()
        {
            //for test
            Console.WriteLine($"[Test] End Of Sequence");

            QuitExcel();
            QuitDriver();
        }
        private void QuitExcel()
        {
            m_workbook.Close();
            RemoveObject(m_workbook);

            m_application.Quit();
            RemoveObject(m_application);
        }
        private void QuitDriver()
        {
            m_driver.Close();
            m_driver.Quit();
        }
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                EndSelenium();
            }
            catch { }
        }

        #region ENUM
        enum CheckType { NetError, CantFindInputForm, ServiceUnavailable, BadGateway }
        #endregion



        #region METHOD : MAIN
        private void InitFormName()
        {
            this.Text = m_versionName;
        }
        private void InitMembers()
        {
            if (m_driver == null) m_driver = new ChromeDriver();
            if (m_dialogClass_list == null) m_dialogClass_list = new List<DialogClass>();
            if (m_random == null) m_random = new Random();
            SetExceptWordArray();
        }
        private void InitInputBox()
        {
            input_ExceptText.Text = m_exceptWord;
            input_RangeOfColumn.Text = m_default_startColumn.ToString();
            input_StartRow.Text = m_default_startRow.ToString();
            input_EndRow.Text = m_defualt_endRow.ToString();
            input_sheetName.Text = m_sheetName;
        }
        private void InitExcelApplication()
        {
            if (m_filePath != "")
            {
                if (m_application == null) m_application = new Microsoft.Office.Interop.Excel.Application();
                m_workbook = m_application.Workbooks.Open(Filename: m_filePath);
            }
        }
        private void ShowEndMessage()
        {
            MessageBox.Show($"[System] 완료 되었습니다.\n - 시간 : {DateTime.Now - m_startTime}\n - 위치 : {m_filePath + " - " + m_resultPath}");
        }
        private void FinishAndRefreshUI()
        {
            List<Control> list_constrols = new List<Control>();
            list_constrols.Add(lb_ExcelPath);
            list_constrols.Add(input_ExcelPath);
            list_constrols.Add(btn_BrowseExcel);
            list_constrols.Add(lb_sheetName);
            list_constrols.Add(input_sheetName);
            list_constrols.Add(lb_ParsingRange);
            list_constrols.Add(lb_Column);
            list_constrols.Add(input_RangeOfColumn);
            list_constrols.Add(lb_RowRange);
            list_constrols.Add(input_StartRow);
            list_constrols.Add(lb_RowRange);
            list_constrols.Add(input_EndRow);
            list_constrols.Add(lb_ExceptText);
            list_constrols.Add(input_ExceptText);
            list_constrols.Add(btn_LoadExcel);
            list_constrols.Add(lb_workComplete);
            list_constrols.Add(lb_resultPath);
            list_constrols.Add(input_resultBox);


            foreach (Control c in list_constrols)
            {
                c.Enabled = !c.Enabled;
                c.Visible = !c.Visible;
            }
            input_resultBox.Text = m_resultFinalPath;
        }
        #endregion



        #region METHOD : ON INPUT BUTTON
        private void btn_BrowseExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog OFD = new OpenFileDialog();
            if (OFD.ShowDialog() == DialogResult.OK)
            {
                input_ExcelPath.Clear();
                input_ExcelPath.Text = OFD.FileName;
                m_filePath = OFD.FileName;
            }
        }
        private void btn_LoadExcel_Click(object sender, EventArgs e) //★Main Method
        {
            if (m_filePath != "")
            {
                InitMembers();
                InitExcelApplication();

                m_startTime = DateTime.Now;

                Worksheet worksheet = m_workbook.Worksheets.get_Item(m_sheetName) as Worksheet;
                if (worksheet == null)
                {
                    MessageBox.Show($"[System : error!!!] WorkSheet({m_sheetName})을 불러오지 못했습니다.");
                    return;
                }

                m_application.Visible = true;
                Range usedRange = worksheet.UsedRange;
                //for test
                Console.WriteLine("[Test] UsedRange 의 마지막 Row 인덱스? = " + usedRange.Rows.Count);
                try
                {
                    for (int i = m_default_startRow; i <= usedRange.Rows.Count; ++i)
                    {
                        if (m_defualt_endRow != -1 && i > m_defualt_endRow) break; //입력한 End Row 까지 도착하여 탈출

                        string classID = null;
                        try
                        {
                            classID = worksheet.Cells[i, 1].Value2.ToString();
                        }
                        catch
                        { }


                        string classDesc = null;
                        try { classDesc = GetStringWithoutExceptWord((worksheet.Cells[i, m_default_startColumn] as Range)); }
                        catch { classDesc = worksheet.Cells[i, m_default_startColumn].Value2.ToString(); }

                        if (classDesc == null || classDesc == "")
                        {
                            Console.WriteLine($"[System : exception!] 행({i}) 값이 비어있으므로 다음으로 Continue");
                            continue;
                        }
                        classDesc = worksheet.Cells[i, m_default_startColumn].Value2.ToString();

                        Console.WriteLine($"ClassID({classID}) / ClassDesc({classDesc})");

                        DialogClass dialogClass = new DialogClass(
                            id: classID,
                            desc: classDesc //제외 단어를 삭제한 형태
                            );

                        m_dialogClass_list.Add(dialogClass);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"[System : exception!!!]\n{ex.Message}\n{ex.StackTrace}");
                }

                try
                {
                    DoSelenium();
                }
                catch { }
                finally
                {
                    SetResultSheet();
                    RemoveObject(worksheet);
                }
            }
            EndSelenium();
            ShowEndMessage();
            FinishAndRefreshUI();
        }
        #endregion



        #region METHOD : ON INPUT TEXT
        private void input_ExcelPath_TextChanged(object sender, EventArgs e)
        {

        }
        private void input_StartRow_TextChanged(object sender, EventArgs e)
        {
            if (IsInputString(input_StartRow.Text))
            {
                m_default_startRow = GetIntByString(input_StartRow.Text);
            }
            else
            {
                MessageBox.Show($"입력값({input_StartRow.Text})은 숫자가 아닙니다. 정수를 입력하세요.");
                input_StartRow.Text = m_default_startRow.ToString();
            }
        }
        private void input_EndRow_TextChanged(object sender, EventArgs e)
        {
            if (IsInputString(input_EndRow.Text))
            {
                m_defualt_endRow = GetIntByString(input_EndRow.Text);
            }
            else
            {
                MessageBox.Show($"입력값({input_EndRow.Text})은 숫자가 아닙니다. 정수를 입력하세요.");
                input_EndRow.Text = m_defualt_endRow.ToString();
            }
        }
        private bool IsInputString(string input)
        {
            int tempInt = GetIntByString(input);
            if (tempInt == default) return false;
            return true;
        }
        private int GetIntByString(string input)
        {
            return int.Parse(input);
        }
        private void input_sheetName_TextChanged(object sender, EventArgs e)
        {
            if (input_sheetName.Text == "" || input_sheetName.Text == null)
            {
                MessageBox.Show($"[System : error!!!] sheet name 에 정상적인 값이 입력되지 않았습니다.");
                input_sheetName.Text = m_sheetName;
                return;
            }

            m_sheetName = input_sheetName.Text;
        }
        #endregion


        [System.Serializable]
        public class DialogClass
        {
            public string m_ID;
            public string m_description;
            public bool m_needCheck { get { return m_result_list.Count == 0 ? false : true; } }
            public bool m_isParsed = false;
            public List<ResultClass> m_result_list = new List<ResultClass>();

            public DialogClass(string id, string desc)
            {
                m_ID = id;
                m_description = desc;
                m_result_list = new List<ResultClass>();
            }
        }



        [System.Serializable]
        public class ResultClass
        {
            public string m_targetString;
            public string m_checkString;
            public string m_help;

            public ResultClass(string targetString, string checkString, string help)
            {
                m_targetString = targetString;
                m_checkString = checkString;
                m_help = help;
            }
        }
    }
}
