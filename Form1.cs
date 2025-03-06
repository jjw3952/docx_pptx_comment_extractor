using System;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Drawing;
using System.Text;

class CommentsExtractor : Form
{
    private Button selectFileButton;
    private RichTextBox outputBox;
    private string[] selectedFiles;
    private ToolTip toolTip;

    public CommentsExtractor()
    {
        // Initialize the form
        this.Text = "Select DOCX/PPTX files from which to extract comments";
        this.Size = new System.Drawing.Size(800, 400);

        selectFileButton = new Button
        {
            Text = "Click here to select files.",
            Dock = DockStyle.Top,
            BackColor = System.Drawing.Color.Yellow,
            Font = new Font(this.Font.FontFamily, 14),
            Height = 50
        };
        selectFileButton.Click += SelectFiles;

        outputBox = new RichTextBox
        {
            Multiline = true,
            Dock = DockStyle.Fill,
            ScrollBars = RichTextBoxScrollBars.Vertical,
            Font = new Font(this.Font.FontFamily, 10)
        };

        Controls.Add(outputBox);
        Controls.Add(selectFileButton);

        outputBox.SelectionFont = new Font(outputBox.Font, FontStyle.Bold);
        outputBox.AppendText(
            "Select a single PPTX file, or single/multiple" +
            "DOCX files from which to extract comments." +
            Environment.NewLine + Environment.NewLine);

        // Initialize the ToolTip and set the text for the selectFileButton
        toolTip = new ToolTip();
        toolTip.SetToolTip(selectFileButton, "Select a single PPTX file, or \r\nsingle multiple DOCX files from which to extract comments.");
    }

    private void SelectFiles(object sender, EventArgs e)
    {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Filter = "Word & PowerPoint Files (*.docx, *.pptx)|*.docx;*.pptx";
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFiles = openFileDialog.FileNames;
                if (ValidateFileSelection(selectedFiles))
                {
                    outputBox.AppendText("Selected Files: " + string.Join(", ", selectedFiles) + Environment.NewLine);
                    ExtractComments();
                }
                else
                {
                    MessageBox.Show("Please select either all DOCX files or a single PPTX file, not a combination of both or multiple PPTX files.", "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }

    private bool ValidateFileSelection(string[] files)
    {
        bool hasDocx = files.Any(f => f.EndsWith(".docx"));
        bool hasPptx = files.Any(f => f.EndsWith(".pptx"));
        bool multiplePptx = files.Count(f => f.EndsWith(".pptx")) > 1;
        return !(hasDocx && hasPptx) && !multiplePptx;
    }

    private void ExtractComments()
    {
        using (SaveFileDialog saveFileDialog = new SaveFileDialog())
        {
            saveFileDialog.Filter = "CSV Files|*.csv";
            saveFileDialog.FileName = "comments.csv"; // Set the default file name
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string savedFile = saveFileDialog.FileName;
                try
                {
                    if (selectedFiles[0].EndsWith(".pptx"))
                    {
                        ExtractCommentsFromPptx(selectedFiles[0], savedFile);
                        outputBox.AppendText(Environment.NewLine + "Saved File: " + savedFile + Environment.NewLine + Environment.NewLine);
                        MessageBox.Show("Comments extracted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        DataTable commentsTable = new DataTable();
                        commentsTable.Columns.Add("CommentID");
                        commentsTable.Columns.Add("Page"); // Page column for DOCX files
                        commentsTable.Columns.Add("Comment");
                        commentsTable.Columns.Add("Reviewer");
                        commentsTable.Columns.Add("Date");
                        commentsTable.Columns.Add("relFile"); // Add relFile column
                        commentsTable.Columns.Add("File");
                        foreach (var file in selectedFiles)
                        {
                            ExtractCommentsFromDocx(file, commentsTable);
                        }

                        // Remove the temporary relFile column before writing to CSV
                        commentsTable.Columns.Remove("relFile");

                        WriteToCsv(commentsTable, savedFile);
                        outputBox.AppendText(Environment.NewLine + "Saved File: " + savedFile + Environment.NewLine + Environment.NewLine);
                        MessageBox.Show("Comments extracted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (IOException ex)
                {
                    MessageBox.Show("The file is currently in use and cannot be overwritten. Please close the file and try again.", "File In Use", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    outputBox.AppendText(Environment.NewLine + "Error: " + ex.Message + Environment.NewLine + Environment.NewLine);
                }
            }
        }
    }

    public static void ExtractCommentsFromPptx(string file, string writeto)
    {
        string copyPath = "pptx_copy.zip";
        File.Copy(file, copyPath, true);

        string extractPath = "pptx_copy";
        if (Directory.Exists(extractPath))
            Directory.Delete(extractPath, true);
        ZipFile.ExtractToDirectory(copyPath, extractPath);

        string authorFile = Path.Combine(extractPath, "ppt/commentAuthors.xml");
        bool isLegacyFormat = false;
        if (!File.Exists(authorFile))
        {
            authorFile = Path.Combine(extractPath, "ppt/authors.xml");
            isLegacyFormat = true;
        }

        XDocument authorDoc = XDocument.Load(authorFile);
        var authors = authorDoc.Descendants().Where(e => e.Name.LocalName == (isLegacyFormat ? "author" : "cmAuthor"))
            .Select(e => new { Id = e.Attribute("id")?.Value, Name = e.Attribute("name")?.Value })
            .ToList();

        DataTable commentsTable = new DataTable();
        commentsTable.Columns.Add("CommentID");
        commentsTable.Columns.Add("Slide", typeof(int)); // Ensure the Slide column is of type int
        commentsTable.Columns.Add("Comment");
        commentsTable.Columns.Add("Reviewer");
        commentsTable.Columns.Add("Date");
        commentsTable.Columns.Add("relFile");
        commentsTable.Columns.Add("File");

        int commentId = 1; // Initialize CommentID
        string originalFileName = Path.GetFileName(file); // Get the original file name

        string commentsPath = Path.Combine(extractPath, "ppt/comments");
        if (Directory.Exists(commentsPath))
        {
            foreach (var commentFile in Directory.GetFiles(commentsPath))
            {
                XDocument commentDoc = XDocument.Load(commentFile);
                var comments = commentDoc.Descendants().Where(e => e.Name.LocalName == (isLegacyFormat ? "cm" : "cm"))
                    .Select(e => new
                    {
                        Id = e.Attribute("authorId")?.Value,
                        Date = e.Attribute("dt")?.Value ?? e.Attribute("created")?.Value, // Handle legacy format date
                        Text = e.Descendants().FirstOrDefault(te => te.Name.LocalName == (isLegacyFormat ? "t" : "text"))?.Value
                    }).ToList();

                foreach (var comment in comments)
                {
                    string reviewer = authors.FirstOrDefault(a => a.Id == comment.Id)?.Name ?? "Unknown";
                    string formattedDate = DateTime.TryParse(comment.Date, out DateTime dateValue) ? dateValue.ToString("MM/dd/yyyy") : comment.Date;
                    commentsTable.Rows.Add(commentId++, 0, comment.Text, reviewer, formattedDate, Path.GetFileName(commentFile), originalFileName); // Add CommentID and formatted date
                }
            }
        }

        string relsPath = Path.Combine(extractPath, "ppt/slides/_rels");
        if (Directory.Exists(relsPath))
        {
            foreach (var relFile in Directory.GetFiles(relsPath))
            {
                XDocument relDoc = XDocument.Load(relFile);
                var commentTargets = relDoc.Descendants().Where(e => e.Name.LocalName == "Relationship" &&
                                                                      e.Attribute("Type")?.Value.Contains("comments") == true)
                    .Select(e => e.Attribute("Target")?.Value.Replace("../comments/", ""))
                    .ToList();

                if (commentTargets.Any())
                {
                    int slideNumber = int.Parse(Path.GetFileNameWithoutExtension(relFile).Replace("slide", "").Replace(".xml", "").Trim());
                    foreach (var target in commentTargets)
                    {
                        foreach (DataRow row in commentsTable.Rows)
                        {
                            if (row["relFile"].ToString() == target)
                            {
                                row["Slide"] = slideNumber;
                            }
                        }
                    }
                }
            }
        }

        // Remove the temporary relFile column
        commentsTable.Columns.Remove("relFile");

        // Remove rows with non-integer Slide values
        for (int i = commentsTable.Rows.Count - 1; i >= 0; i--)
        {
            if (!int.TryParse(commentsTable.Rows[i]["Slide"].ToString(), out _))
            {
                commentsTable.Rows.RemoveAt(i);
            }
        }

        // Sort commentsTable by Slide column
        commentsTable.DefaultView.Sort = "Slide ASC";
        commentsTable = commentsTable.DefaultView.ToTable();

        // Renumber CommentID based on the sorted order
        int newCommentId = 1;
        foreach (DataRow row in commentsTable.Rows)
        {
            row["CommentID"] = newCommentId++;
        }

        Directory.Delete(extractPath, true);
        File.Delete(copyPath);

        WriteToCsv(commentsTable, writeto);
    }

    public static void ExtractCommentsFromDocx(string file, DataTable commentsTable)
    {
        string copyPath = "docx_copy.zip";
        File.Copy(file, copyPath, true);

        string extractPath = "docx_copy";
        if (Directory.Exists(extractPath))
            Directory.Delete(extractPath, true);
        ZipFile.ExtractToDirectory(copyPath, extractPath);

        string commentsFile = Path.Combine(extractPath, "word/comments.xml");
        if (!File.Exists(commentsFile))
        {
            MessageBox.Show("No comments found in the DOCX file.", "No Comments", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        XDocument commentsDoc = XDocument.Load(commentsFile);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var comments = commentsDoc.Descendants().Where(e => e.Name.LocalName == "comment")
            .Select(e => new
            {
                Id = e.Attribute("id")?.Value,
                Author = e.Attribute(XName.Get("author", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"))?.Value,
                Date = e.Attribute(XName.Get("date", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"))?.Value,
                Text = string.Concat(e.Descendants(w + "t").Select(t => (string)t))//e.Descendants().FirstOrDefault(te => te.Name.LocalName == "t")?.Value
            }).ToList();

        int commentId = 1; // Initialize CommentID
        string originalFileName = Path.GetFileName(file); // Get the original file name

        foreach (var comment in comments)
        {
            string formattedDate = DateTime.TryParse(comment.Date, out DateTime dateValue) ? dateValue.ToString("MM/dd/yyyy") : comment.Date;
            commentsTable.Rows.Add(commentId++, "", comment.Text, comment.Author, formattedDate, "", originalFileName); // Add CommentID and formatted date, set relFile to blank
        }

        //    //// Gather all comments from the 'Comments' column
        //    string allComments = "";
        //    foreach (DataRow row in commentsTable.Rows)
        //    {
        //        allComments += row["Comment"].ToString() + Environment.NewLine;
        //    }

        //    // Show all comments in a MessageBox
        //    MessageBox.Show(allComments, "All Comments");

        //    Directory.Delete(extractPath, true);
        //    File.Delete(copyPath);
    }
    
    private static void WriteToCsv(DataTable table, string filePath)
    {
        try
        {
            using (StreamWriter writer = new StreamWriter(filePath, false, new UTF8Encoding(true)))
            {
                writer.WriteLine(string.Join(",", table.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
                foreach (DataRow row in table.Rows)
                {
                    //var fields = row.ItemArray.Select(field =>
                    //{
                    //    string fieldString = field.ToString();
                    //    if (fieldString.Contains(",") || fieldString.Contains("\"") || fieldString.Contains("'"))
                    //    {
                    //        fieldString = "\"" + fieldString.Replace("\"", "\"\"") + "\"";
                    //    }
                    //    return fieldString;
                    //});
                    var fields = row.ItemArray.Select(field =>
                    {
                        string fieldString = field.ToString();

                        // Normalize curly quotes to standard quotes
                        fieldString = fieldString.Replace("“", "\"").Replace("”", "\"")
                                                 .Replace("‘", "\'").Replace("’", "\'");
                        //.Replace("…", "...");

                        // Check if field contains special characters that require quoting
                        if (fieldString.Contains(",") || fieldString.Contains("\"") || fieldString.Contains("\'") || fieldString.Contains("\n"))
                        {
                            fieldString = "\"" + fieldString.Replace("\"", "\"\"") + "\"";
                        }

                        return fieldString;
                    });
                    writer.WriteLine(string.Join(",", fields));
                }
            }
        }
        catch (IOException ex)
        {
            throw new IOException("Failed to write to the file. It may be in use by another process.", ex);
        }
    }

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new CommentsExtractor());
    }
}

