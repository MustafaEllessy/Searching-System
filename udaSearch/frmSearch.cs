using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using IFilterTextReader;
using IronOcr;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using OfficeOpenXml;
using System.Diagnostics;
using System.Text;
//supported extensions are .txt, .docx,.doc,.pptx,.pdf,.xlsx
namespace udaSearch
{
    public partial class frmSearch : Form
    {
        public frmSearch()
        {
            InitializeComponent();
        }
        FolderBrowserDialog fbd;
        OpenFileDialog ofd;
        string directoryPath;
        int occurances;
        Dictionary<string,string> matchingTxtFiles;
       // Dictionary<string, string> matchingRTFFiles;
        Dictionary<string, string> matchingWordFiles;
        Dictionary<string, string> matchingPdfFiles;
        Dictionary<string, string> matchingPowerPointFiles;
        Dictionary<string, string> matchingExcelFiles;
        List<string> matchingTxtFilesList;
     //   List<string> matchingRTFFilesList;
        List<string> matchingWordFilesList;
        List<string> matchingPdfFilesList;
        List<string> matchingPowerPointFilesList;
        List<string> matchingExcelFilesList;
        private void frmSearch_Load(object sender, EventArgs e)
        {

          
            fbd = new FolderBrowserDialog();
            ofd = new OpenFileDialog();
            directoryPath = "";
            occurances = 0;
            matchingTxtFiles = new Dictionary<string, string>();
           // matchingRTFFiles = new Dictionary<string, string>();
            matchingWordFiles = new Dictionary<string, string>();
            matchingPdfFiles = new Dictionary<string, string>();
            matchingExcelFiles = new Dictionary<string, string>();
            matchingPowerPointFiles = new Dictionary<string, string>();
            matchingTxtFilesList= new List<string>();
           // matchingRTFFilesList = new List<string>();
            matchingWordFilesList = new List<string>();
            matchingPdfFilesList = new List<string>();
            matchingPowerPointFilesList = new List<string>();
            matchingExcelFilesList = new List<string>();
        }

        private void btnChooseDirectory_Click(object sender, EventArgs e)
        {
            DialogResult dr = fbd.ShowDialog();
            if (dr == DialogResult.OK)
            {
                directoryPath = fbd.SelectedPath;
            }
        }

        private void btnStartSearching_Click(object sender, EventArgs e)
        {
       
        }

        private void Clear()
        {
            totalFiles = 0;
            matchingTxtFiles.Clear();
           // matchingRTFFiles.Clear();
            matchingWordFiles.Clear();
            matchingPdfFiles.Clear();
            matchingExcelFiles.Clear();
            matchingPowerPointFiles.Clear();
            matchingTxtFilesList.Clear();
           // matchingRTFFilesList.Clear();
            matchingWordFilesList.Clear();
            matchingPdfFilesList.Clear();
            matchingExcelFilesList.Clear();
            matchingPowerPointFilesList.Clear();
            dgvSearch.Invoke((MethodInvoker)delegate () { dgvSearch.Rows.Clear(); });
            lblTxt.Invoke(new MethodInvoker(delegate { lblTxt.Text = 0 + ""; }));
            lblWord.Invoke(new MethodInvoker(delegate { lblWord.Text = 0 + ""; }));
            lblPDF.Invoke(new MethodInvoker(delegate { lblPDF.Text = 0 + ""; }));
            lblPPT.Invoke(new MethodInvoker(delegate { lblPPT.Text = 0 + ""; }));
            lblExcel.Invoke(new MethodInvoker(delegate { lblExcel.Text = 0 + ""; }));
            lblTotal.Invoke(new MethodInvoker(delegate { lblTotal.Text = 0 + ""; }));
        }



        private void StartSearching(string dPath)
        {
            try
            {

            if (!Directory.Exists(dPath))
            {
                return;
            }
            string[] existingFiles = Directory.GetFiles(dPath);
            int searchIndex = -1;
            cbxExt.Invoke((MethodInvoker)delegate {  searchIndex = cbxExt.SelectedIndex; });
           
            

            for (int i = 0; i < existingFiles.Count(); i++)
            {
                string fileExt=( System.IO.Path.GetExtension(existingFiles[i])).ToLower();
                if(searchIndex==1 && fileExt!=".txt")
                    {
                        continue;
                    }
                    else if ((searchIndex == 2 && fileExt != ".doc") && (searchIndex == 2 && fileExt != ".docx"))
                    {
                        continue;
                    }
                    else if ((searchIndex == 3 && fileExt != ".pdf") && (searchIndex == 3 && fileExt != ".jpg") && (searchIndex == 3 && fileExt != ".jpeg") && (searchIndex == 3 && fileExt != ".png"))
                    {
                        continue;

                    }
                    else if (searchIndex == 4 && fileExt != ".pptx")
                    {
                        continue ;
                    }
                    else if (searchIndex == 5 && fileExt != ".xlsx")
                    {
                        continue;
                    }
                   
                    SearchThisFile(existingFiles[i]);
            }
            string[] existingDirectories = Directory.GetDirectories(dPath);
            for (int i = 0; i < existingDirectories.Count(); i++)
            {
                StartSearching(existingDirectories[i]);
            }
        }
             catch (Exception)
            {

            }
        }
        int totalFiles = 0;
        private void SearchThisFile(string fileName)
        {
            try
            {
               
            string fileExt = ( System.IO.Path.GetExtension(fileName)).ToLower();
            if (fileExt == ".txt")
            {
                using (var fs = new FileStream(fileName, FileMode.Open))
                {
                    var canRead = fs.CanRead;
                    if (!canRead) return;
                }
                SearchTXT(fileName);
            }
            //else if (fileExt == ".rtf")
            //{
            //    SearchRTF(fileName);
            //}
            else if (fileExt == ".docx" || fileExt == ".doc")
            {
                using (var fs = new FileStream(fileName, FileMode.Open))
                {
                    var canRead = fs.CanRead;
                    if (!canRead) return;
                }
                SearchDOCX(fileName);
            }
            else if (fileExt == ".pdf" || fileExt == ".png" || fileExt == ".jpg" || fileExt == ".jpeg" )
            {
                using (var fs = new FileStream(fileName, FileMode.Open))
                {
                    var canRead = fs.CanRead;
                    if (!canRead) return;
                }
                SearchPDF(fileName);
            }
            else if (fileExt == ".pptx")
            {
                using (var fs = new FileStream(fileName, FileMode.Open))
                {
                    var canRead = fs.CanRead;
                    if (!canRead) return;
                }
                SearchPPTX(fileName);
            }
            else if (fileExt == ".xlsx" )
            {
                using (var fs = new FileStream(fileName, FileMode.Open))
                {
                    var canRead = fs.CanRead;
                    if (!canRead) return;
                }
                SearchXLSX(fileName);
            }
                totalFiles++;
                lblTotal.Invoke((MethodInvoker)delegate { lblTotal.Text = totalFiles + ""; });
            }
            catch (Exception)
            {

            }
        }

        private void SearchXLSX(string fileName)
        {
            try
            {

            if (new FileInfo(fileName).Length == 0)
            {
                return;
            }
            using (ExcelPackage xlPackage = new ExcelPackage())
            {
                var sb = new StringBuilder(); //this is your data
                using (var stream = File.OpenRead(fileName))
                {
                    xlPackage.Load(stream);
                }
                foreach (var workSheet in xlPackage.Workbook.Worksheets)
                {
                    var myWorksheet = workSheet;//select sheet here
                    if (myWorksheet.Cells.Count() == 0) continue;
                    var totalRows = myWorksheet.Dimension.End.Row;
                    var totalColumns = myWorksheet.Dimension.End.Column;
                    for (int rowNum = 1; rowNum <= totalRows; rowNum++) //select starting row here
                    {
                        var row = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                        sb.AppendLine(string.Join(",", row));
                    }
                }

                string fileData = sb.ToString();
                    string searchWord = "";
                    txtSearch.Invoke((MethodInvoker)delegate { searchWord = txtSearch.Text.ToLower(); });
                    string fileNameWithExt = System.IO.Path.GetFileName(fileName);
                if (isIN(fileNameWithExt, fileData, searchWord))
                {
                    matchingExcelFiles.Add(fileNameWithExt, fileName);
                //    matchingExcelFilesList.Add(fileNameWithExt);
                        fillInTable(fileName);
                    }
            }
            }
            catch (Exception)
            {

            }

        }

        private void SearchPPTX(string fileName)
        {
            try
            {

            string file = fileName;
            if (new FileInfo(fileName).Length == 0)
            {
                return;
            }
            string fileData = "";
            int numberOfSlides = CountSlides(file);
            for (int i = 0; i < numberOfSlides; i++)
            {
                string slideData = GetAllTextInSlide(fileName, i);
                fileData += slideData;
            }
                string searchWord = "";
                txtSearch.Invoke((MethodInvoker)delegate { searchWord = txtSearch.Text.ToLower(); });
                string fileNameWithExt = System.IO.Path.GetFileName(fileName);
            if (isIN(fileNameWithExt, fileData, searchWord))
            {
                matchingPowerPointFiles.Add(fileNameWithExt, fileName);
              //  matchingPowerPointFilesList.Add(fileNameWithExt);
                     fillInTable( fileName);

            }
            }
            catch (Exception)
            {

            }
        }
        public static int CountSlides(string presentationFile)
        {
            // Open the presentation as read-only.
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
            {
                // Pass the presentation to the next CountSlides method
                // and return the slide count.
                return CountSlides(presentationDocument);
            }
        }
        // Count the slides in the presentation.
        public static int CountSlides(PresentationDocument presentationDocument)
        {
            // Check for a null document object.
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            int slidesCount = 0;

            // Get the presentation part of document.
            PresentationPart presentationPart = presentationDocument.PresentationPart;
            // Get the slide count from the SlideParts.
            if (presentationPart != null)
            {
                slidesCount = presentationPart.SlideParts.Count();
            }
            // Return the slide count to the previous method.
            return slidesCount;
        }
        public static string GetAllTextInSlide(string presentationFile, int slideIndex)
        {
            // Open the presentation as read-only.
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
            {
                // Pass the presentation and the slide index
                // to the next GetAllTextInSlide method, and
                // then return the array of strings it returns. 
                return GetAllTextInSlide(presentationDocument, slideIndex);
            }
        }
        public static string GetAllTextInSlide(PresentationDocument presentationDocument, int slideIndex)
        {
            // Verify that the presentation document exists.
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            // Verify that the slide index is not out of range.
            if (slideIndex < 0)
            {
                throw new ArgumentOutOfRangeException("slideIndex");
            }

            // Get the presentation part of the presentation document.
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation part and presentation exist.
            if (presentationPart != null && presentationPart.Presentation != null)
            {
                // Get the Presentation object from the presentation part.
                Presentation presentation = presentationPart.Presentation;

                // Verify that the slide ID list exists.
                if (presentation.SlideIdList != null)
                {
                    // Get the collection of slide IDs from the slide ID list.
                    DocumentFormat.OpenXml.OpenXmlElementList slideIds =
                        presentation.SlideIdList.ChildElements;

                    // If the slide ID is in range...
                    if (slideIndex < slideIds.Count)
                    {
                        // Get the relationship ID of the slide.
                        string slidePartRelationshipId = (slideIds[slideIndex] as SlideId).RelationshipId;

                        // Get the specified slide part from the relationship ID.
                        SlidePart slidePart =
                            (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);

                        // Pass the slide part to the next method, and
                        // then return the array of strings that method
                        // returns to the previous method.
                        return GetAllTextInSlide(slidePart);
                    }
                }
            }

            // Else, return null.
            return null;
        }
        public static string GetAllTextInSlide(SlidePart slidePart)
        {
            // Verify that the slide part exists.
            if (slidePart == null)
            {
                throw new ArgumentNullException("slidePart");
            }

            // Create a new linked list of strings.
            LinkedList<string> texts = new LinkedList<string>();

            // If the slide exists...
            if (slidePart.Slide != null)
            {
                // Iterate through all the paragraphs in the slide.
                foreach (DocumentFormat.OpenXml.Drawing.Paragraph paragraph in
                    slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
                {
                    // Create a new string builder.                    
                    StringBuilder paragraphText = new StringBuilder();

                    // Iterate through the lines of the paragraph.
                    foreach (DocumentFormat.OpenXml.Drawing.Text text in
                        paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                    {
                        // Append each line to the previous lines.
                        paragraphText.Append(text.Text);
                    }

                    if (paragraphText.Length > 0)
                    {
                        // Add each paragraph to the linked list.
                        texts.AddLast(paragraphText.ToString());
                    }
                }
            }

            if (texts.Count > 0)
            {
                string data = "";
                string[] dataArr = texts.ToArray();
                // Return an array of strings.
                for (int i = 0; i < dataArr.Count(); i++)
                {
                    data += dataArr[i];
                }
                return data;
            }
            else
            {
                return "";
            }
        }
        private void SearchPDF(string fileName)
        {
            try
            {

            StringBuilder text = new StringBuilder();
            //bool isEnglish = hasEnglishLetters(txtSearch.Text.ToLower());
            if (new FileInfo(fileName).Length == 0)
            {
                return;
            }

                /*   using (PdfReader reader = new PdfReader(fileName))
                   {
                       for (int i = 1; i <= reader.NumberOfPages; i++)
                       {
                           text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                       }
                   }
                */
                var Ocr = new IronTesseract();
                Ocr.Language = OcrLanguage.Arabic;
                Ocr.AddSecondaryLanguage(OcrLanguage.English);
                using (var Input = new OcrInput(fileName))
                {
                    var Result = Ocr.Read(Input);
                    string currentSubText = Result.Text;
                    text.Append( currentSubText) ;
                }


                string searchWord = "";
                txtSearch.Invoke((MethodInvoker)delegate { searchWord=txtSearch.Text.ToLower(); });
                string fileData = text.ToString();
            string fileNameWithExt = System.IO.Path.GetFileName(fileName);
            if (isIN(fileNameWithExt, fileData, searchWord))
            {
                matchingPdfFiles.Add(fileNameWithExt, fileName);
                //matchingPdfFilesList.Add(fileNameWithExt);
                    fillInTable(fileName);
                }
            }
            catch (Exception)
            {

            }
        }

        private bool hasEnglishLetters(string text)
        {
            bool isEnglish = false;
            for (int i = 0; i < text.Length; i++)
            {
                if (isEnglishLetter(text[i]))
                {
                    isEnglish = true;
                    break;
                }
            }
            return isEnglish; 
        }

        private bool isEnglishLetter(char v)
        {
            if(v=='a' || v == 'b' || v == 'c' || v == 'd' || v == 'e' || v == 'f' || v == 'g' || v == 'h' || v == 'i' || v == 'j' || v == 'k' || v == 'l' || v == 'm' || v == 'm' || v == 'n' || v == 'o' || v == 'p' || v == 'q' || v == 'r' || v == 's' || v == 't' || v == 'u' || v == 'v' || v == 'w' ||  v == 'x' || v == 'y' || v == 'z' )
            {
                return true;
            }
            return false;
        }

        private void SearchDOCX(string fileName)
        {
            try
            {
                string searchWord = "";
                txtSearch.Invoke((MethodInvoker)delegate { searchWord = txtSearch.Text.ToLower(); });

                if (new FileInfo(fileName).Length == 0)
            {
                return;
            }
            TextReader reader = new FilterReader(fileName);
            string txt = "";
            using (reader)
            {
                txt = reader.ReadToEnd();
            }
            string fileData = txt;
            string fileNameWithExt = System.IO.Path.GetFileName(fileName);
            if (isIN(fileNameWithExt, fileData, searchWord))
            {
                matchingWordFiles.Add(fileNameWithExt, fileName);
               // matchingWordFilesList.Add(fileNameWithExt);
                    fillInTable(fileName);
                }
        }
              catch (Exception)
            {

            }
        }
        private void SearchRTF(string fileName)
        {
            string searchWord = "";
            txtSearch.Invoke((MethodInvoker)delegate { searchWord = txtSearch.Text.ToLower(); }); ;

            string fileData = File.ReadAllText(fileName,Encoding.UTF8);
            MessageBox.Show(fileData);
            string fileNameWithExt = System.IO.Path.GetFileName(fileName);
            if (isIN(fileNameWithExt, fileData, searchWord))
            {
                //matchingRTFFiles.Add(fileNameWithExt, fileName);
                //matchingRTFFilesList.Add(fileNameWithExt);
                fillInTable(fileName);
            }
        }

        private void SearchTXT(string fileName)
        {
            try
            {

                string searchWord = "";
                txtSearch.Invoke((MethodInvoker)delegate { searchWord = txtSearch.Text.ToLower(); });
                string fileData = File.ReadAllText(fileName);
            if (new FileInfo(fileName).Length == 0)
            {
                return;
            }
            string fileNameWithExt = System.IO.Path.GetFileName(fileName);
          
            if (isIN(fileNameWithExt, fileData, searchWord))
            {
                     matchingTxtFiles.Add(fileNameWithExt, fileName);
                    // matchingTxtFilesList.Add(fileNameWithExt);
                    fillInTable( fileName);
            }
            }
            catch (Exception)
            {

            }
        }

        private void fillInTable( string fileName)
        {
            int lastRowIndex = dgvSearch.RowCount-1;
            int rowsCount = dgvSearch.Rows.Count;
            string fileNameWithExt = System.IO.Path.GetFileName(fileName);
            string fileExt= System.IO.Path.GetExtension(fileName);
            dgvSearch.Invoke((MethodInvoker)delegate {
                if (fileExt == ".txt")
                {
                   if(rowsCount==0|| dgvSearch.Rows[lastRowIndex].Cells[0].Value!="")
                    {
                        object[] obj = new object[] { fileNameWithExt, "", "", "", "" };
                        dgvSearch.Rows.Add(obj);
                    }
                   else
                    {
                        int rowIndex = -1;
                        for (int i = 0; i < rowsCount; i++)
                        {
                            string firstPlace = dgvSearch.Rows[i].Cells[0].Value.ToString();
                            if(firstPlace=="")
                            {
                                rowIndex = i;
                                break;
                            }    
                        }
                        if(rowIndex==-1)
                        {
                            MessageBox.Show("Logic Error");
                            return;
                        }
                        dgvSearch.Rows[rowIndex].Cells[0].Value= fileNameWithExt;

                    }
                    lblTxt.Invoke(new MethodInvoker(delegate { lblTxt.Text = (Convert.ToInt32(lblTxt.Text) + 1) + ""; }));
             

                }
                else if (fileExt == ".doc" || fileExt == ".docx")
                {

                    if (rowsCount == 0 || dgvSearch.Rows[lastRowIndex].Cells[1].Value != "")
                    {
                        object[] obj = new object[] { "", fileNameWithExt, "", "", "" };
                        dgvSearch.Rows.Add(obj);

                    }
                    else
                    {
                        int rowIndex = -1;
                        for (int i = 0; i < rowsCount; i++)
                        {
                            string firstPlace = dgvSearch.Rows[i].Cells[1].Value.ToString();
                            if (firstPlace == "")
                            {
                                rowIndex = i;
                                break;
                            }
                        }
                        if (rowIndex == -1)
                        {
                            MessageBox.Show("Logic Error");
                            return;
                        }
                        dgvSearch.Rows[rowIndex].Cells[1].Value = fileNameWithExt;

                    }

                    lblWord.Invoke(new MethodInvoker(delegate { lblWord.Text = (Convert.ToInt32(lblWord.Text) + 1) + ""; }));
         
                }
                else if (fileExt == ".xlsx")
                {
                    if (rowsCount == 0 || dgvSearch.Rows[lastRowIndex].Cells[4].Value != "")
                    {
                        object[] obj = new object[] { "", "", "", "", fileNameWithExt };
                        dgvSearch.Rows.Add(obj);

                    }
                    else
                    {
                        int rowIndex = -1;
                        for (int i = 0; i < rowsCount; i++)
                        {
                            string firstPlace = dgvSearch.Rows[0].Cells[4].Value.ToString();
                            if (firstPlace == "")
                            {
                                rowIndex = i;
                                break;
                            }
                        }
                        if (rowIndex == -1)
                        {
                            MessageBox.Show("Logic Error");
                            return;
                        }
                        dgvSearch.Rows[rowIndex].Cells[4].Value = fileNameWithExt;

                    }
                    lblExcel.Invoke(new MethodInvoker(delegate { lblExcel.Text = (Convert.ToInt32(lblExcel.Text) + 1) + ""; }));

                }
                else if (fileExt == ".pdf")
                {
                    if (rowsCount == 0 || dgvSearch.Rows[lastRowIndex].Cells[2].Value != "")
                    {
                        object[] obj = new object[] { "", "", fileNameWithExt, "", "" };
                        dgvSearch.Rows.Add(obj);

                    }
                    else
                    {
                        int rowIndex = -1;
                        for (int i = 0; i < rowsCount; i++)
                        {
                            string firstPlace = dgvSearch.Rows[i].Cells[2].Value.ToString();
                            if (firstPlace == "")
                            {
                                rowIndex = i;
                                break;
                            }
                        }
                        if (rowIndex == -1)
                        {
                            MessageBox.Show("Logic Error");
                            return;
                        }
                        dgvSearch.Rows[rowIndex].Cells[2].Value = fileNameWithExt;

                    }

                    lblPDF.Invoke(new MethodInvoker(delegate { lblPDF.Text = (Convert.ToInt32(lblPDF.Text) + 1) + ""; }));

                }
                else if (fileExt == ".pptx")
                {
                    if (rowsCount == 0 || dgvSearch.Rows[lastRowIndex].Cells[3].Value != "")
                    {
                        object[] obj = new object[] { "", "", "", fileNameWithExt, "" };
                        dgvSearch.Rows.Add(obj);

                    }
                    else
                    {
                        int rowIndex = -1;
                        for (int i = 0; i < rowsCount; i++)
                        {
                            string firstPlace = dgvSearch.Rows[i].Cells[3].Value.ToString();
                            if (firstPlace == "")
                            {
                                rowIndex = i;
                                break;
                            }
                        }
                        if (rowIndex == -1)
                        {
                            MessageBox.Show("Logic Error");
                            return;
                        }
                        dgvSearch.Rows[rowIndex].Cells[3].Value = fileNameWithExt;

                    }
                    lblPPT.Invoke(new MethodInvoker(delegate { lblPPT.Text = (Convert.ToInt32(lblPPT.Text) + 1) + ""; }));

                }
            });
    
        }

        private bool isIN(string fName, string data, string word)
        {
            int charsCount = data.Length - word.Length;
            occurances = 0;
            for (int i = 0; i <= charsCount; i++)
            {
                string wordToTest = data.Substring(i, word.Length);
                if (word == wordToTest)
                {
                    return true;
                }

            }
            return false;   
        }


        private void btnMatchingFiles_Click(object sender, EventArgs e)
        {
            ShowFolders();
        }

        private void ShowFolders()
        {
           
        }

        private void Search()
        {
            StartSearching(directoryPath);
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
        }

        private void FillTable()
        {
            dgvSearch.Invoke((MethodInvoker)delegate {
                int maxRows = -1;
                if (matchingExcelFilesList.Count > maxRows) maxRows = matchingExcelFilesList.Count;
                if (matchingPdfFilesList.Count > maxRows) maxRows = matchingPdfFilesList.Count;
                if (matchingTxtFilesList.Count > maxRows) maxRows = matchingTxtFilesList.Count;
                if (matchingWordFilesList.Count > maxRows) maxRows = matchingWordFilesList.Count;
                //  if (matchingRTFFilesList.Count > maxRows) maxRows = matchingRTFFilesList.Count;
                if (matchingPowerPointFilesList.Count > maxRows) maxRows = matchingPowerPointFilesList.Count;
                for (int i = 0; i < maxRows; i++)
                {
                    object[] obj = new object[] { "", "", "", "", "" };
                    if (matchingExcelFilesList.Count > i)
                    {
                        obj[4] = matchingExcelFilesList[i];
                    }
                    if (matchingPdfFilesList.Count > i)
                    {
                        obj[2] = matchingPdfFilesList[i];
                    }
                    if (matchingPowerPointFilesList.Count > i)
                    {
                        obj[3] = matchingPowerPointFilesList[i];
                    }
                    //if (matchingRTFFilesList.Count > i)
                    //{
                    //    obj[1] = matchingRTFFilesList[i];
                    //}
                    if (matchingTxtFilesList.Count > i)
                    {
                        obj[0] = matchingTxtFilesList[i];
                    }
                    if (matchingWordFilesList.Count > i)
                    {
                        obj[1] = matchingWordFilesList[i];
                    }
                    dgvSearch.Rows.Add(obj);
                }
            });
            lblTxt.Invoke(new MethodInvoker(delegate { lblTxt.Text = matchingTxtFilesList.Count + ""; }));
            lblWord.Invoke(new MethodInvoker(delegate { lblWord.Text = matchingWordFilesList.Count + ""; }));
            lblPDF.Invoke(new MethodInvoker(delegate { lblPDF.Text = matchingPdfFilesList.Count + ""; }));
            lblPPT.Invoke(new MethodInvoker(delegate {lblPPT.Text = matchingPowerPointFilesList.Count + "";}));
            lblExcel.Invoke(new MethodInvoker(delegate { lblExcel.Text = matchingExcelFilesList.Count + ""; }));
            lblTotal.Invoke(new MethodInvoker(delegate { lblTotal.Text = totalFiles + ""; }));
        }

        private void dgvSearch_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            string content = dgvSearch.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            if(content.Length>0)
            {
                string path = "";
               if(e.ColumnIndex==0)
                {
                    path = matchingTxtFiles[content];
                }
                //else if (e.ColumnIndex == 1)
                //{
                //    path = matchingRTFFiles[content];
                //}
                else if (e.ColumnIndex == 1)
                {
                    path = matchingWordFiles[content];
                }
                else if (e.ColumnIndex == 2)
                {
                    path = matchingPdfFiles[content];
                }
                else if (e.ColumnIndex == 3)
                {
                    path = matchingPowerPointFiles[content];
                }
                else if (e.ColumnIndex == 4)
                {
                    path = matchingExcelFiles[content];
                }
               if(path.Length>0)
                {
                    var p = new Process();
                    p.StartInfo = new ProcessStartInfo(path)
                    {
                        UseShellExecute = true
                    };
                    p.Start();
                }
            }
        }

        private void dgvSearch_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void txtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) btnSearch.PerformClick();
        }
        Thread searchTask;
        
        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (fbd.SelectedPath=="")
            {
                MessageBox.Show("Select folder");
                return;
            }
            if(cbxExt.SelectedIndex==-1)
            {
                MessageBox.Show("Select Search Option");
                return;
            }
            searchTask=new Thread(SearchProcess);
            searchTask.Start();
        }

        private void SearchProcess()
        {

            statusStrip1.Invoke((MethodInvoker)delegate {  toolStripStatusLabel2.Text = "Processing.."; });
            
            Clear();
            string searchWord = "";
            txtSearch.Invoke((MethodInvoker)delegate { searchWord = txtSearch.Text.ToLower(); });
            if (searchWord == "") return;

            if (fbd.SelectedPath == "") return;
            Search();
         //   FillTable();
            statusStrip1.Invoke((MethodInvoker)delegate { toolStripStatusLabel2.Text = "Finished"; });

          
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            searchTask.Abort();
        }

        private void cbxExt_SelectedIndexChanged(object sender, EventArgs e)
        {
          
         
        }

        private void frmSearch_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}