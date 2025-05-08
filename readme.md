using System;
//using OfficeOpenXml;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Status;
using System.Runtime.InteropServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Data.OleDb;
using System.Data.Common;
using System.Threading;
using System.Runtime.InteropServices.ComTypes;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using StudentReport;

namespace C0_Correction_Form
{

public partial class OMR : Form
    {
string connectionString = GlobalSettings.ConnectionString;
        DataTable dataTable = new DataTable();
        SqlDataAdapter dataAdapter;
private string Username;
public OMR(string userName)
        {

InitializeComponent();
            Username = userName;
            label12.Text = $"Logged in as: {Username}"; ;

            this.WindowState = FormWindowState.Maximized;

            // Bind text changed event once during initialization
            textBox5.TextChanged += TxtFilter_TextChanged;
            textBox1.TextChanged += TxtFilter_TextChanged;
            textBox2.TextChanged += TxtFilter_TextChanged;
            textBox6.TextChanged += TxtFilter_TextChanged;
            textBox7.TextChanged += TxtFilter_TextChanged;
            textBox9.TextChanged += TxtFilter_TextChanged;
            textBox10.TextChanged += TxtFilter_TextChanged;

            // Load form asynchronously
            this.Load += async (sender, e) => await Form1_LoadAsync(sender, e);
            this.FormClosing += MainForm_FormClosing;
        }
private async Task Form1_LoadAsync(object sender, EventArgs e)
        {
            // Load data and initialize grid view asynchronously
          //  await InitializeGridViewAsync();
FilterData(); // Optionally, call after initial data load
            //dataGridView1_CellClick(null, null);
PopulateComboBox1();
PopulateComboBox2();
        }
private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
dataGridView1.EndEdit(); // End cell editing when the form is closing
        }
private async Task InitializeGridViewAsync()
        {
try
            {
using (SqlConnection connection = new SqlConnection(connectionString))
                {
await connection.OpenAsync();
string query = "SELECT * FROM PaperInfo";

using (SqlCommand command = new SqlCommand(query, connection))
                    {
dataAdapter = new SqlDataAdapter(command);
dataTable = new DataTable();
await Task.Run(() => dataAdapter.Fill(dataTable));
                        dataGridView1.DataSource = dataTable;
if (dataGridView1.Columns.Contains("MARKS"))
                        {
dataGridView1.Columns["MARKS"].ReadOnly = true;
                        }
else
                        {
MessageBox.Show("Column 'MARKS' not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
catch (Exception ex)
            {
MessageBox.Show($"Error initializing data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
dataGridView1.Columns["MARKS"].ReadOnly = true;
        }
private void TxtFilter_TextChanged(object sender, EventArgs e)
        {
FilterData();
            textBox8.Text = dataTable.Rows.Count.ToString();
        }

private void FilterData()
        {
string selectedFilename = string.Empty;

this.Invoke(new Action(() =>
            {
selectedFilename = comboBox1.SelectedItem?.ToString();
            }));
string rollNo = textBox5.Text;
string course = textBox1.Text;
string barcode = textBox2.Text;
string paperCode = textBox6.Text;
string series = textBox7.Text;
string year = textBox9.Text;
string marks = textBox10.Text;


bool anyFilterFilled = !string.IsNullOrEmpty(rollNo) ||
                       !string.IsNullOrEmpty(course) ||
                       !string.IsNullOrEmpty(barcode) ||
                       !string.IsNullOrEmpty(paperCode) ||
                       !string.IsNullOrEmpty(series) ||
                       !string.IsNullOrEmpty(year) ||
marks != null ||
                       !string.IsNullOrEmpty(selectedFilename);

if (anyFilterFilled)
            {
string query = @"SELECT TOP 10000 ID, [ROLL NO], BARCODE, BOOKLET, SERIES, [PAPER CODE], [CENTER CODE], [EXAM TYPE], SEM, COURSE, YEAR, MARKS, filename 
                         FROM PaperInfo 
                         WHERE 1=1";

                List<SqlParameter> parameters = new List<SqlParameter>();

if (!string.IsNullOrEmpty(selectedFilename))
                {
query += " AND filename LIKE @Filename";
parameters.Add(new SqlParameter("@Filename", "%" + selectedFilename + "%"));
                }

if (!string.IsNullOrEmpty(rollNo))
                {
query += " AND [ROLL NO] LIKE @RollNo";
parameters.Add(new SqlParameter("@RollNo", "%" + rollNo + "%"));
                }

if (!string.IsNullOrEmpty(course))
                {
query += " AND COURSE LIKE @Course";
parameters.Add(new SqlParameter("@Course", "%" + course + "%"));
                }

if (!string.IsNullOrEmpty(barcode))
                {
query += " AND barcode LIKE @barcode";
parameters.Add(new SqlParameter("@barcode", "%" + barcode + "%"));
                }

if (!string.IsNullOrEmpty(paperCode))
                {
query += " AND [PAPER CODE] LIKE @PaperCode";
parameters.Add(new SqlParameter("@PaperCode", "%" + paperCode + "%"));
                }

if (!string.IsNullOrEmpty(series))
                {
query += " AND SERIES LIKE @Series";
parameters.Add(new SqlParameter("@Series", "%" + series + "%"));
                }

if (!string.IsNullOrEmpty(year))
                {
query += " AND YEAR LIKE @Year";
parameters.Add(new SqlParameter("@Year", "%" + year + "%"));
                }
if (!string.IsNullOrEmpty(marks))
                {
query += " AND marks = @Marks";
parameters.Add(new SqlParameter("@Marks", marks));
                }

                // Run the query asynchronously
using (SqlConnection connection = new SqlConnection(connectionString))
                {
connection.Open();
using (SqlCommand command = new SqlCommand(query, connection))
                    {
command.Parameters.AddRange(parameters.ToArray());

using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
dataTable = new DataTable();
adapter.Fill(dataTable);
                        }
                    }
                }

this.Invoke(new Action(() =>
                {
                    dataGridView1.DataSource = dataTable;

dataGridView1.Columns["ID"].Visible = false;
dataGridView1.Columns["COURSE"].Visible = false;
dataGridView1.Columns["YEAR"].Visible = false;
dataGridView1.Columns["filename"].Visible = false;

                    dataGridView1.Visible = true;

                    textBox8.Text = dataTable.Rows.Count.ToString(); // Update count display
                }));
            }
else
            {
this.Invoke(new Action(() =>
                {
                    dataGridView1.Visible = false;
                    textBox8.Text = "0";
                }));
            }
        }
//=====================================================================================================CLEAR BUTTON OPERATION================================================================================
private void Clear_button_Click(object sender, EventArgs e)
        {
            textBox10.Text = "";
            textBox1.Text = "";
            textBox5.Text = "";
            textBox2.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox4.Text = "";
if (comboBox1.SelectedItem != null)
            {
                comboBox1.SelectedIndex = -1; // Reset the selection
            }

            dataGridView1.DataSource = null; // Clear any data binding
dataGridView1.Rows.Clear();      // Clear all rows if not data-bound
dataGridView1.Refresh();
            dataGridView1.Visible = false;
            pictureBox1.Visible = false;
        }
private void UpdateRecord(DataRow row, string remark)
        {
string query = "UPDATE PaperInfo SET [ROLL NO] = @RollNo, BARCODE = @Barcode, BOOKLET = @Booklet, SERIES = @Series, [PAPER CODE] = @PaperCode, [CENTER CODE] = @CenterCode, SEM = @Sem, [EXAM TYPE] = @ExamType, remarks = @Remarks WHERE ID = @ID";

using (SqlConnection connection = new SqlConnection(connectionString))
            {
using (SqlCommand command = new SqlCommand(query, connection))
                {
command.Parameters.AddWithValue("@RollNo", row["ROLL NO"]);
command.Parameters.AddWithValue("@Barcode", row["BARCODE"]);
command.Parameters.AddWithValue("@Booklet", row["BOOKLET"]);
command.Parameters.AddWithValue("@Series", row["SERIES"]);
command.Parameters.AddWithValue("@PaperCode", row["PAPER CODE"]);
command.Parameters.AddWithValue("@CenterCode", row["CENTER CODE"]);
command.Parameters.AddWithValue("@Sem", row["SEM"]);
command.Parameters.AddWithValue("@ExamType", row["EXAM TYPE"]);
command.Parameters.AddWithValue("@Remarks", remark);
command.Parameters.AddWithValue("@ID", row["ID"]);

connection.Open();
command.ExecuteNonQuery();
                }
            }
        }
private void button1_Click(object sender, EventArgs e)
        {
dataGridView1.EndEdit();

            DataTable dt = (DataTable)dataGridView1.DataSource;

string remark = textBox4.Text;

bool updatesMade = false;

foreach (DataRow row in dt.Rows)
            {
if (row.RowState == DataRowState.Modified)
                {
UpdateRecord(row, remark);
updatesMade = true;
                }
            }

if (updatesMade)
            {
MessageBox.Show("Changes updated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
else
            {
MessageBox.Show("No changes to update.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
private string GetImageDirectory(DataGridViewCellEventArgs e)
        {
            //string connectionString = ConfigurationManager.ConnectionStrings["con"].ConnectionString;
string imageDirectory = null;

long id = Convert.ToInt64(dataGridView1.Rows[e.RowIndex].Cells["ID"].Value);
string filename = dataGridView1.Rows[e.RowIndex].Cells["FILENAME"].Value.ToString();

string query = "SELECT filename FROM PaperInfo WHERE ID = @ID AND FILENAME = @Filename";

using (SqlConnection connection = new SqlConnection(connectionString))
            {
try
                {
connection.Open();

using (SqlCommand command = new SqlCommand(query, connection))
                    {
command.Parameters.AddWithValue("@ID", id);
command.Parameters.AddWithValue("@Filename", filename);

object result = command.ExecuteScalar();

if (result != null)
                        {
imageDirectory = result.ToString();

                            // Remove the .csv or .xls extension if it exists
imageDirectory = Path.GetFileNameWithoutExtension(imageDirectory);
                        }
else
                        {
MessageBox.Show("No image directory found for the specified ID and filename.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
catch (Exception ex)
                {
MessageBox.Show("An error occurred while retrieving the image directory: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

return imageDirectory;
        }

private string GetNCS_HeaderValueFromDatabase(long id)
        {
string ncsHeaderValue = "";
string selectedFilename = null;

            // Safely access comboBox1.SelectedItem on the UI thread
if (comboBox1.InvokeRequired)
            {
comboBox1.Invoke(new Action(() =>
                {
if (comboBox1.SelectedItem != null)
                    {
selectedFilename = comboBox1.SelectedItem.ToString();
                    }
                }));
            }
else
            {
if (comboBox1.SelectedItem != null)
                {
selectedFilename = comboBox1.SelectedItem.ToString();
                }
            }

            //string query = "SELECT CASE WHEN CHARINDEX('.', [NCS Header]) > 0 THEN [NCS Header] ELSE SUBSTRING([NCS Header], 4, 6) END FROM PaperInfo WHERE ID = @ID";
string query = "SELECT BARCODE FROM PaperInfo WHERE ID = @ID";
if (!string.IsNullOrEmpty(selectedFilename))
            {
query += " AND filename = @Filename";
            }

using (SqlConnection connection = new SqlConnection(connectionString))
            {
using (SqlCommand command = new SqlCommand(query, connection))
                {
command.Parameters.AddWithValue("@ID", id);  // Correctly passing ID as bigint

if (!string.IsNullOrEmpty(selectedFilename))
                    {
command.Parameters.AddWithValue("@Filename", selectedFilename);
                    }

connection.Open();

object result = command.ExecuteScalar();

if (result != null)
                    {
ncsHeaderValue = result.ToString();
                    }
                }
            }

return ncsHeaderValue;
        }
private void DisplayImage(long id, string filename)
        {
string ncsHeaderValue = GetNCS_HeaderValueFromDatabase(id);  // Since filename is the clicked value

if (filename.EndsWith(".csv") || filename.EndsWith(".xls"))
            {
filename = filename.Remove(filename.Length - 4);
            }
else if (filename.EndsWith(".xlsx"))
            {
filename = filename.Remove(filename.Length - 5);
            }

string imageDirectory;
if (ncsHeaderValue.Contains(".jpg"))
            {
imageDirectory = Path.Combine(filename, $"{ncsHeaderValue}");
            }
else
            {
imageDirectory = Path.Combine(filename, $"{ncsHeaderValue}.jpg");
            }

if (File.Exists(imageDirectory))
            {
try
                {
if (pictureBox1.Image != null)
                    {
pictureBox1.Image.Dispose();
                    }

                    pictureBox1.Image = Image.FromFile(imageDirectory);
                }
catch (Exception ex)
                {
MessageBox.Show("Error loading image: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
else
            {
MessageBox.Show("Image not found: " + imageDirectory, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



private string ncsHeaderValue;

private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true; // Prevent default Enter behavior

if (dataGridView1.CurrentCell != null)
                {
int rowIndex = dataGridView1.CurrentCell.RowIndex;
int columnIndex = dataGridView1.CurrentCell.ColumnIndex;

                    // Handle interaction with the current cell
                    DataGridViewCellEventArgs cellEventArgs = new DataGridViewCellEventArgs(columnIndex, rowIndex);
HandleCellInteraction(cellEventArgs);

                    // Find the next visible row
int nextRowIndex = rowIndex + 1;
while (nextRowIndex < dataGridView1.Rows.Count && !dataGridView1.Rows[nextRowIndex].Visible)
                    {
nextRowIndex++;
                    }

                    // Ensure we don't exceed row limits
if (nextRowIndex < dataGridView1.Rows.Count && dataGridView1.Rows[nextRowIndex].Visible)
                    {
int firstVisibleColumnIndex = GetFirstVisibleColumnIndex();

                        // Scroll to the next row
                        dataGridView1.FirstDisplayedScrollingRowIndex = nextRowIndex;

                        // Move to the new row and first visible column
                        dataGridView1.CurrentCell = dataGridView1.Rows[nextRowIndex].Cells[firstVisibleColumnIndex];
                    }
                }
            }
        }
private void ShowImage(int rowIndex, int columnIndex)
        {
if (rowIndex >= 0 && columnIndex >= 0)
            {
                // Retrieve the ID and filename from the DataGridView row
long id = Convert.ToInt64(dataGridView1.Rows[rowIndex].Cells["ID"].Value);  // Assuming "ID" column exists
string filename = dataGridView1.Rows[rowIndex].Cells["FILENAME"].Value?.ToString();

if (!string.IsNullOrEmpty(filename))
                {
DisplayImage(id, filename);
                }
            }
        }

private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
ShowImage(e.RowIndex, e.ColumnIndex);
            }
        }
private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
HandleCellInteraction(e);
        }
private int GetFirstVisibleColumnIndex()
        {
foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
if (col.Visible)
                {
return col.Index;
                }
            }
return 0;
        }
private void HandleCellInteraction(DataGridViewCellEventArgs e)
        {
            pictureBox1.Visible = true;
if (e == null || e.RowIndex < 0 || e.ColumnIndex < 0)
            {
return;
            }
try
            {
                DataGridViewRow selectedRow = dataGridView1.Rows[e.RowIndex];
string idValue = selectedRow.Cells["ID"].Value?.ToString();
string paperCode = selectedRow.Cells["PAPER CODE"].Value?.ToString();
string rollNo = selectedRow.Cells["ROLL NO"].Value?.ToString();
string filename = selectedRow.Cells["FILENAME"].Value?.ToString();
string series = selectedRow.Cells["SERIES"].Value?.ToString();

if (string.IsNullOrWhiteSpace(idValue) || !long.TryParse(idValue, out long id))
                {
MessageBox.Show("Invalid ID value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
return;
                }
Task.Run(() => DisplayImage(id, filename));

UpdateMarks(id, paperCode, series, e.RowIndex);
            }
catch (Exception ex)
            {
MessageBox.Show($"An error occured: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            //finally
            //{
            //    pictureBox1.Visible =false;
            //}
        }

private void UpdateMarks(long id, string paperCode, string series, int rowIndex)
        {
int maxMarks;
            List<string> ansKeyList = new List<string>();
            List<string> studentAnsList = new List<string>();

try
            {
using (SqlConnection connection = new SqlConnection(connectionString))
                {
connection.Open();
maxMarks = GetMaxMarks(connection, paperCode, series);
ansKeyList = GetAnswerKey(connection, paperCode, series);
studentAnsList = GetStudentAnswers(connection, id, paperCode, series, ansKeyList.Count);
int obtainedMarks = CalculateMarks(ansKeyList, studentAnsList, maxMarks);
UpdateMarksInDatabase(connection, id, paperCode, series, obtainedMarks);
UpdateGridView(rowIndex, obtainedMarks);
                }
            }
catch (Exception ex)
            {
                //MessageBox.Show($"Error updating marks: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
private int GetMaxMarks(SqlConnection connection, string paperCode, string series)
        {
string query = "SELECT MAX(MAX_MARKS) FROM ANSKEY WHERE KEY_TYPE = @PaperCode and SET_NO = @Series";
using (SqlCommand command = new SqlCommand(query, connection))
            {
command.Parameters.AddWithValue("@PaperCode", paperCode);
command.Parameters.AddWithValue("@Series", series);
object result = command.ExecuteScalar();
return result != null && result != DBNull.Value ? Convert.ToInt32(result) : 0;
            }
        }
private List<string> GetAnswerKey(SqlConnection connection, string paperCode, string series)
        {
string query = "SELECT ANS_NO FROM ANSKEY WHERE KEY_TYPE = @PaperCode AND SET_NO = @Series ORDER BY QUES";
            List<string> answerKeyList = new List<string>();
using (SqlCommand command = new SqlCommand(query, connection))
            {
command.Parameters.AddWithValue("@PaperCode", paperCode);
command.Parameters.AddWithValue("@Series", series);
using (SqlDataReader reader = command.ExecuteReader())
                {
while (reader.Read())
                    {
answerKeyList.Add(reader["ANS_NO"].ToString());
                    }
                }

            }
return answerKeyList;
        }
private List<string> GetStudentAnswers(SqlConnection connection, long id, string paperCode, string series, int questionCount)
        {
            string query = $"SELECT {string.Join(", ", Enumerable.Range(1, questionCount).Select(i => $"ANS{i}"))} " +
                   "FROM PaperInfo WHERE ID = @Id AND [PAPER CODE] = @PaperCode AND SERIES = @Series";
            List<string> studentAnsList = new List<string>();
using (SqlCommand command = new SqlCommand(query, connection))
            {
command.Parameters.AddWithValue("@Id", id);
command.Parameters.AddWithValue("@PaperCode", paperCode);
command.Parameters.AddWithValue("@Series", series);
using (SqlDataReader reader = command.ExecuteReader())
                {
if (reader.Read())
                    {
for (int i = 1; i <= questionCount; i++)
                        {
studentAnsList.Add(reader[$"ANS{i}"].ToString());
                        }
                    }
                }
            }
return studentAnsList;
        }
private int CalculateMarks(List<string> ansKeyList, List<string> studentAnsList, int maxMarks)
        {
int correctCount = ansKeyList.Zip(studentAnsList, (key, ans) => key == ans).Count(match => match);
return (int)(double)correctCount * maxMarks;
        }
private void UpdateMarksInDatabase(SqlConnection connection, long id, string paperCode, string series, int marks)
        {
string query = "UPDATE PaperInfo SET MARKS = @Marks WHERE ID = @Id AND [PAPER CODE] = @PaperCode AND SERIES = @Series";
using (SqlCommand command = new SqlCommand(query, connection))
            {
command.Parameters.AddWithValue("@Marks", marks);
command.Parameters.AddWithValue("@Id", id);
command.Parameters.AddWithValue("@PaperCode", paperCode);
command.Parameters.AddWithValue("@Series", series);
command.ExecuteNonQuery();
            }
        }
private void UpdateGridView(int rowIndex, int marks)
        {
dataGridView1.Rows[rowIndex].Cells["MARKS"].Value = marks;
        }
private void PopulateComboBox1()
        {
            //string connectionString = ConfigurationManager.ConnectionStrings["con"].ConnectionString;
string query = "SELECT DISTINCT FILENAME FROM PAPERINFO WHERE (MARKS IS NULL OR MARKS = 0)";
            comboBox1.Enabled = true;
            comboBox1.DropDownStyle = ComboBoxStyle.DropDown;
using (SqlConnection connection = new SqlConnection(connectionString))
            {
try
                {
connection.Open();
using (SqlCommand command = new SqlCommand(query, connection))
                    {
using (SqlDataReader reader = command.ExecuteReader())
                        {
while (reader.Read())
                            {
comboBox1.Items.Add(reader["filename"].ToString());
                            }
                        }
                    }
                }
catch (Exception ex)
                {
MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
if (comboBox1.SelectedItem != null)
            {
string selectedFilename = comboBox1.SelectedItem.ToString();
PopulateGridView1(selectedFilename);
            }
        }
private void PopulateGridView1(string filename)
        {
            //string connectionString = ConfigurationManager.ConnectionStrings["con"].ConnectionString;
string query = "SELECT ID, [ROLL NO], BOOKLET, BARCODE, SERIES, [PAPER CODE], [CENTER CODE], [EXAM TYPE], SEM, MARKS, FILENAME  FROM PaperInfo WHERE filename = @filename";

            DataTable dataTable = new DataTable();
using (SqlConnection connection = new SqlConnection(connectionString))
            {
try
                {
connection.Open();
using (SqlCommand command = new SqlCommand(query, connection))
                    {
command.Parameters.AddWithValue("@filename", filename);

using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
adapter.Fill(dataTable);
                            textBox8.Text = dataTable.Rows.Count.ToString();
                        }
                    }
                }
catch (Exception ex)
                {
MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            dataGridView1.DataSource = dataTable;
            dataGridView1.Visible = true;
dataGridView1.Columns["ID"].Visible = false;
dataGridView1.Columns["filename"].Visible = false;
        }

//=======================================================================================UPLOAD FILE==================================================================================================================
private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Filter = "Excel and CSV Files|*.xls;*.xlsx;*.csv";
            openFileDialog1.Title = "Select File";

if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
try
                {
string fileName = openFileDialog1.FileName;
                    textBox3.Text = fileName;
ImportFileToDatabase(fileName);
                }
catch (Exception ex)
                {
MessageBox.Show("Error: " + ex.Message);
                }
            }
        }


private void ImportFileToDatabase(string fileName)
        {
            //string connectionString = ConfigurationManager.ConnectionStrings["con"].ConnectionString;
string fullFilePath = textBox3.Text;
if (CheckIfFileIsImported(fullFilePath))
            {
MessageBox.Show("The file has already been imported.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
return;
            }
            DataTable dataTable = new DataTable();
using (StreamReader reader = new StreamReader(fileName))
            {
string[] columns = reader.ReadLine().Split(',');

foreach (string column in columns)
                {
string trimmedColumnName = column.Trim();
dataTable.Columns.Add(trimmedColumnName); // Add columns without specifying data type
                }

                // Ensure mandatory columns are present
if (!dataTable.Columns.Contains("ROLL NO"))
                {
MessageBox.Show("The file does not contain the required column 'ROLL NO'.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
return;
                }
if (!dataTable.Columns.Contains("PAPER CODE"))
                {
MessageBox.Show("The file does not contain the required column 'PAPER CODE'.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
return;
                }

if (!dataTable.Columns.Contains("FILENAME"))
                {
dataTable.Columns.Add("FILENAME");
                }
if (!dataTable.Columns.Contains("DATE_N_TIME"))
                {
dataTable.Columns.Add("DATE_N_TIME", typeof(DateTime));
                }

while (!reader.EndOfStream)
                {
string[] values = reader.ReadLine().Split(',');

                    DataRow newRow = dataTable.NewRow();

for (int i = 0; i < values.Length; i++)
                    {
if (i < columns.Length && dataTable.Columns.Contains(columns[i]))
                        {
newRow[columns[i]] = values[i];
                        }
                    }
newRow["FILENAME"] = fullFilePath;
newRow["DATE_N_TIME"] = DateTime.Now;
dataTable.Rows.Add(newRow);
                }
            }

using (SqlConnection connection = new SqlConnection(connectionString))
            {
try
                {
connection.Open();

using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                    {
                        bulkCopy.DestinationTableName = "PaperInfo";

foreach (DataColumn column in dataTable.Columns)
                        {
bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName);
                        }
bulkCopy.WriteToServer(dataTable);
                    }

                    string updateQuery1 = @"UPDATE PaperInfo SET [ROLL NO] = REPLACE([ROLL NO], '""','') WHERE [ROLL NO] LIKE '%""%'";
string updateQuery2 = @"";

string initialCatalog = connection.Database;
string filename = textBox3.Text;

if(initialCatalog == "KVP_2024")
                    {
                        updateQuery2 = @"UPDATE Paperinfo SET [PAPER CODE] = '24/' + [PAPER CODE]
                                        WHERE [PAPER CODE] NOT LIKE '24/%' and filename = @filename";
                    }
if (initialCatalog == "KVP_2025")
                    {
                        updateQuery2 = @"UPDATE Paperinfo SET [PAPER CODE] = '25/' + [PAPER CODE]
                                        WHERE [PAPER CODE] NOT LIKE '25/%' and filename = @filename";
                    }

using (SqlCommand command = new SqlCommand(updateQuery1, connection))
                    {
command.Parameters.AddWithValue("@filename", filename);
command.ExecuteNonQuery();
                    }
using (SqlCommand command = new SqlCommand(updateQuery2, connection))
                    {
command.Parameters.AddWithValue("@filename", filename);
command.ExecuteNonQuery();
                    }
MessageBox.Show("File imported successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
catch (Exception ex)
                {
MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

private bool CheckIfFileIsImported(string fileWithoutExtension)
        {
            //string connectionString = ConfigurationManager.ConnectionStrings["con"].ConnectionString;
bool isFileImported = false;

string query = "SELECT COUNT(*) FROM PaperInfo WHERE filename = @filename";

using (SqlConnection connection = new SqlConnection(connectionString))
            {
try
                {
connection.Open();

using (SqlCommand command = new SqlCommand(query, connection))
                    {
command.Parameters.AddWithValue("@filename", fileWithoutExtension);

int count = (int)command.ExecuteScalar();
isFileImported = count > 0;
                    }
                }
catch (Exception ex)
                {
MessageBox.Show("An error occurred while checking if the file is already imported: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

return isFileImported;
        }
private void PopulateGridView(string fileName)
        {
string query = "SELECT ID, [ROLL NO], BOOKLET, BARCODE, SERIES, [PAPER CODE], [CENTER CODE], [EXAM TYPE], SEM, MARKS, FILENAME  FROM PaperInfo WHERE filename = @filename";
using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlDataAdapter adapter = new SqlDataAdapter(query, conn);
adapter.SelectCommand.Parameters.AddWithValue("@filename", fileName);
                DataTable dt = new DataTable();
adapter.Fill(dt);

                dataGridView1.DataSource = dt;
            }
        }

public void ExecuteWithRetry(Action action, int maxRetries = 3)
        {
int delay = 1000;
for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
try
                {
action();
return;
                }
catch (SqlException ex) when (ex.Number == 1205) // Deadlock
                {
if (attempt == maxRetries)
throw;

Thread.Sleep(delay);
delay *= 2;
                }
            }
        }

private void button3_Click(object sender, EventArgs e)
        {
if (comboBox1.SelectedItem == null)
            {
MessageBox.Show("Please select a file:", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
return; // Exit the method if no file is selected
            }

string selectedFilename = comboBox1.SelectedItem.ToString();
            List<PaperInfo> paperInfoList = new List<PaperInfo>();

using (SqlConnection connection = new SqlConnection(connectionString))
                {
connection.Open();
string query = "SELECT [ROLL NO], [PAPER CODE], SERIES FROM PaperInfo WHERE filename = @filename";
using (SqlCommand command = new SqlCommand(query, connection))
                    {
command.Parameters.AddWithValue("@filename", selectedFilename);
using (SqlDataReader reader = command.ExecuteReader())
                        {
while (reader.Read())
                            {
paperInfoList.Add(new PaperInfo
                                {
                                    RollNo = reader["ROLL NO"].ToString(),
                                    PaperCode = reader["PAPER CODE"].ToString(),
                                    Series = reader["SERIES"].ToString()
                                });
                            }
                        }
                    }
                }
try
                {
foreach (var paperInfo in paperInfoList)
                    {
CalculateAndUpdateMarks(paperInfo.RollNo, paperInfo.PaperCode, paperInfo.Series, selectedFilename);
                    }
MessageBox.Show("Marks calculated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
PopulateGridView(selectedFilename);
                }
catch (Exception ex)
                {
MessageBox.Show($"An error occurred while calculating marks: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

        }

private void CalculateAndUpdateMarks(string rollNo, string paperCode, string series, string filename)
        {
int maxMarks = 0;

using (SqlConnection connection = new SqlConnection(connectionString))
                {
connection.Open();

string queryMaxMarks = "SELECT MAX(MAX_MARKS) AS MaxMarks FROM ANSKEY WHERE KEY_TYPE = @PaperCode AND SET_NO = @Series";
using (SqlCommand command = new SqlCommand(queryMaxMarks, connection))
                    {
command.Parameters.AddWithValue("@PaperCode", paperCode);
command.Parameters.AddWithValue("@Series", series);
object result = command.ExecuteScalar();
if (result != null && result != DBNull.Value)
                        {
maxMarks = Convert.ToInt32(result);
                        }
                    }

                    List<string> ansNoList = new List<string>();
using (SqlCommand command = new SqlCommand("SELECT ANS_NO FROM ANSKEY WHERE KEY_TYPE = @paperCode AND [SET_NO] = @series ORDER BY Ques", connection))
                    {
command.Parameters.AddWithValue("@paperCode", paperCode);
command.Parameters.AddWithValue("@series", series);

using (SqlDataReader reader = command.ExecuteReader())
                        {
while (reader.Read())
                            {
ansNoList.Add(reader["ANS_NO"].ToString());
                            }
                        }
                    }

                    List<string> ansList = new List<string>();
using (SqlCommand command = new SqlCommand(@"
        SELECT ANS1, ANS2, ANS3, ANS4, ANS5, ANS6, ANS7, ANS8, ANS9, ANS10,
               ANS11, ANS12, ANS13, ANS14, ANS15, ANS16, ANS17, ANS18, ANS19, ANS20,
               ANS21, ANS22, ANS23, ANS24, ANS25, ANS26, ANS27, ANS28, ANS29, ANS30,
               ANS31, ANS32, ANS33, ANS34, ANS35, ANS36, ANS37, ANS38, ANS39, ANS40,
               ANS41, ANS42, ANS43, ANS44, ANS45, ANS46, ANS47, ANS48, ANS49, ANS50,
               ANS51, ANS52, ANS53, ANS54, ANS55, ANS56, ANS57, ANS58, ANS59, ANS60,
               ANS61, ANS62, ANS63, ANS64, ANS65, ANS66, ANS67, ANS68, ANS69, ANS70,
               ANS71, ANS72, ANS73, ANS74, ANS75, ANS76, ANS77, ANS78, ANS79, ANS80,
               ANS81, ANS82, ANS83, ANS84, ANS85, ANS86, ANS87, ANS88, ANS89, ANS90,
               ANS91, ANS92, ANS93, ANS94, ANS95, ANS96, ANS97, ANS98, ANS99, ANS100,
               ANS101, ANS102, ANS103, ANS104, ANS105, ANS106, ANS107, ANS108, ANS109, ANS110,
               ANS111, ANS112, ANS113, ANS114, ANS115, ANS116, ANS117, ANS118, ANS119, ANS120,
               ANS121, ANS122, ANS123, ANS124, ANS125, ANS126, ANS127, ANS128, ANS129, ANS130,
               ANS131, ANS132, ANS133, ANS134, ANS135, ANS136, ANS137, ANS138, ANS139, ANS140,
               ANS141, ANS142, ANS143, ANS144, ANS145, ANS146, ANS147, ANS148, ANS149, ANS150
        FROM PaperInfo WHERE [ROLL NO] = @RollNo AND [PAPER CODE] = @PaperCode AND [SERIES] = @Series AND filename = @filename", connection))
                    {
command.Parameters.AddWithValue("@RollNo", rollNo);
command.Parameters.AddWithValue("@PaperCode", paperCode);
command.Parameters.AddWithValue("@Series", series);
command.Parameters.AddWithValue("@filename", filename);

using (SqlDataReader reader = command.ExecuteReader())
                        {
if (reader.Read())
                            {
for (int i = 1; i <= 150; i++)
                                {
ansList.Add(reader["ANS" + i].ToString());
                                }
                            }
else
                            {
MessageBox.Show("No marks found for the given roll number, paper code, series, and filename.");
return;
                            }
                        }
                    }

int count = 0;
for (int i = 0; i < ansNoList.Count; i++)
                    {
string[] correctAnswers = ansNoList[i].Split(',');
string studentAnswer = ansList[i];
bool anyMatch = correctAnswers.Any(ca => ca.Trim().Equals(studentAnswer, StringComparison.OrdinalIgnoreCase));
                        //bool allMatch = correctAnswers.All(ca => ca.Trim().Equals(studentAnswer, StringComparison.OrdinalIgnoreCase));

if (anyMatch)
                        {
count++;
                        }
                    }

int getMarks = maxMarks * count;
using (SqlCommand updateCommand = new SqlCommand("UPDATE PaperInfo SET MARKS = @Marks WHERE [ROLL NO] = @RollNo AND [PAPER CODE] = @PaperCode AND [SERIES] = @Series AND filename = @filename", connection))
                    {
updateCommand.Parameters.AddWithValue("@Marks", getMarks);
updateCommand.Parameters.AddWithValue("@RollNo", rollNo);
updateCommand.Parameters.AddWithValue("@PaperCode", paperCode);
updateCommand.Parameters.AddWithValue("@Series", series);
updateCommand.Parameters.AddWithValue("@filename", filename);
updateCommand.ExecuteNonQuery();
                    }
                }
        }

public class PaperInfo
        {
public string RollNo { get; set; }
public string PaperCode { get; set; }
public string Series { get; set; }
public string Filename { get; set; }  // Added to store filename
        }
private void button5_Click(object sender, EventArgs e)
        {
if (comboBox1.SelectedItem != null)
            {
string selectedFilename = comboBox1.SelectedItem.ToString();

string newRemarks = textBox4.Text; // Assuming you have a TextBox for inputting the new remarks

string query = "UPDATE paperInfo SET remarks = @remarks WHERE filename = @filename";

                //string connectionString = ConfigurationManager.ConnectionStrings["con"].ConnectionString;

using (SqlConnection connection = new SqlConnection(connectionString))
                {
try
                    {
connection.Open();

using (SqlCommand command = new SqlCommand(query, connection))
                        {
command.Parameters.AddWithValue("@remarks", newRemarks);
command.Parameters.AddWithValue("@filename", selectedFilename);

int rowsAffected = command.ExecuteNonQuery();

if (rowsAffected > 0)
                            {
MessageBox.Show("Remarks updated successfully for " + rowsAffected + " row(s).", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
else
                            {
MessageBox.Show("No records updated. Please check the filename.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
catch (Exception ex)
                    {
MessageBox.Show("An error occurred while updating the remarks: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
else
            {
MessageBox.Show("Please select a filename from the ComboBox.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
private void PopulateComboBox2()
        {
            //string connectionString = ConfigurationManager.ConnectionStrings["con"].ConnectionString;
string query = "SELECT distinct filename FROM paperInfo WHERE remarks IS NOT NULL AND LTRIM(RTRIM(remarks)) <> ''";

            // Clear the existing items in ComboBox2
comboBox2.Items.Clear();

using (SqlConnection connection = new SqlConnection(connectionString))
            {
try
                {
connection.Open();

using (SqlCommand command = new SqlCommand(query, connection))
                    {
using (SqlDataReader reader = command.ExecuteReader())
                        {
while (reader.Read())
                            {
string fullPath = reader["filename"].ToString();

string shortFileName = Path.GetFileName(fullPath);

comboBox2.Items.Add(shortFileName);
                            }
                        }
                    }
                }
catch (Exception ex)
                {
MessageBox.Show("An error occurred while populating ComboBox2: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
private void PaperCode_correction_Click(object sender, EventArgs e)
        {

        }

private void Set_correction_Click(object sender, EventArgs e)
        {

        }

private void RollNo_correction_Click_1(object sender, EventArgs e)
        {
            // Define your SQL query
            string query = @"SELECT p.ID, p.[ROLL NO], p.BARCODE, p.BOOKLET, p.[PAPER CODE], p.SERIES, p.SEM, p.[EXAM TYPE], p.MARKS, p.FILENAME, p.[CENTER CODE] 
                     FROM PaperInfo p 
                     LEFT JOIN reg r ON p.[ROLL NO] = r.ROLL_NO  
                     WHERE r.ROLL_NO IS NULL";

if (!string.IsNullOrEmpty(textBox1.Text))
            {
query += " AND p.course = @Course";
            }

if (!string.IsNullOrEmpty(textBox9.Text))
            {
query += " AND p.year = @Year";
            }


using (SqlConnection con = new SqlConnection(connectionString))
            {
try
                {
                    SqlCommand cmd = new SqlCommand(query, con);
if (textBox1.Text != null)
                    {
cmd.Parameters.AddWithValue("@Course", textBox1.Text);
                    }

if (textBox9.Text != null)
                    {
cmd.Parameters.AddWithValue("@Year", textBox9.Text);
                    }
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
da.Fill(dt);
                    dataGridView1.DataSource = dt;

                    textBox8.Text = dt.Rows.Count.ToString();
                }
catch (Exception ex)
                {
MessageBox.Show("An error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
dataGridView1.Columns["filename"].Visible = false;
dataGridView1.Columns["ID"].Visible = false;
            }
        }
private void duplicate_correction_Click(object sender, EventArgs e)
        {
string query = @"SELECT D.ID, D.[ROLL NO], D.BARCODE, D.BOOKLET, D.SERIES, D.[PAPER CODE],
D.[CENTER CODE], D.[EXAM TYPE], D.SEM, D.MARKS, D.FILENAME
                            FROM 
                            (
                              SELECT [ROLL NO], [PAPER CODE], COUNT(*) AS COUNT
                              FROM PAPERINFO
                              GROUP BY [ROLL NO], [PAPER CODE]
                              HAVING COUNT(*) > 1
                            ) AS T 
                            LEFT JOIN
                            (
                                SELECT ID, [ROLL NO], BARCODE, BOOKLET, SERIES, [PAPER CODE], [CENTER CODE], [EXAM TYPE], SEM, MARKS, FILENAME
                                FROM PAPERINFO
                            ) AS D
                            ON D.[ROLL NO] = T.[ROLL NO]
                            AND D.[PAPER CODE] = T.[PAPER CODE]
                            ORDER BY D.[PAPER CODE]";
using (SqlConnection con = new SqlConnection(connectionString))
            {
try
                {
                    SqlCommand cmd = new SqlCommand(query, con);
if (textBox1.Text != null)
                    {
cmd.Parameters.AddWithValue("@Course", textBox1.Text);
                    }

if (textBox9.Text != null)
                    {
cmd.Parameters.AddWithValue("@Year", textBox9.Text);
                    }
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
da.Fill(dt);
                    dataGridView1.DataSource = dt;

                    textBox8.Text = dt.Rows.Count.ToString();
                }
catch (Exception ex)
                {
MessageBox.Show("An error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
dataGridView1.Columns["filename"].Visible = false;
dataGridView1.Columns["ID"].Visible = false;
            }
        }
private void set_correction_Click_1(object sender, EventArgs e)
        {
            // ISKO BAAD ME BANANA HE 
        }
private void paperCode_correction_Click_1(object sender, EventArgs e)
        {
            string query = @"SELECT D.ID, D.[ROLL NO], D.BARCODE, D.BOOKLET, D.SERIES, D.[PAPER CODE], D.[CENTER CODE], D.[EXAM TYPE], D.SEM, D.MARKS, D.FILENAME FROM PaperInfo D
                            LEFT JOIN STUDENT_MARKS T ON D.[ROLL NO] = T.ROLL_NO AND D.[PAPER CODE] = T.PAPER_CODE WHERE T.PAPER_CODE IS NULL ORDER BY D.[PAPER CODE] DESC";
using (SqlConnection con = new SqlConnection(connectionString))
            {
try
                {
                    SqlCommand cmd = new SqlCommand(query, con);
if (textBox1.Text != null)
                    {
cmd.Parameters.AddWithValue("@Course", textBox1.Text);
                    }

if (textBox9.Text != null)
                    {
cmd.Parameters.AddWithValue("@Year", textBox9.Text);
                    }
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
da.Fill(dt);
                    dataGridView1.DataSource = dt;

                    textBox8.Text = dt.Rows.Count.ToString();
                }
catch (Exception ex)
                {
MessageBox.Show("An error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
dataGridView1.Columns["filename"].Visible = false;
dataGridView1.Columns["ID"].Visible = false;
            }
        }

    }
}


//        private void button3_Click(object sender, EventArgs e)
//        {

//            string selectedFilename = comboBox1.SelectedItem.ToString();
//            List<PaperInfo> paperInfoList = new List<PaperInfo>();
//            ExecuteWithRetry(() =>
//            {
//                using (SqlConnection connection = new SqlConnection(connectionString))
//                {
//                    connection.Open();
//                    string query = "SELECT [ROLL NO], [PAPER CODE], SERIES FROM PaperInfo";
//                    if (!string.IsNullOrEmpty(selectedFilename))
//                    {
//                        query += " WHERE filename = @filename";
//                    }
//                    using (SqlCommand command = new SqlCommand(query, connection))
//                    {
//                        if (!string.IsNullOrEmpty(selectedFilename))
//                        {
//                            command.Parameters.AddWithValue("@filename", selectedFilename);
//                        }
//                        using (SqlDataReader reader = command.ExecuteReader())
//                        {
//                            while (reader.Read())
//                            {
//                                paperInfoList.Add(new PaperInfo
//                                {
//                                    RollNo = reader["ROLL NO"].ToString(),
//                                    PaperCode = reader["PAPER CODE"].ToString(),
//                                    Series = reader["SERIES"].ToString() 
//                                });
//                            }
//                        }
//                    }
//                }

//                Parallel.ForEach(paperInfoList, paperInfo =>
//                {
//                    try
//                    {
//                        CalculateAndUpdateMarks(paperInfo.RollNo, paperInfo.PaperCode, paperInfo.Series, selectedFilename);
//                    }
//                    catch (Exception ex)
//                    {
//                        MessageBox.Show($"An error occurred while calculating marks: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
//                    }
//                });

//                MessageBox.Show("Marks calculated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
//                PopulateGridView(selectedFilename);
//            });
//        }

//        private void CalculateAndUpdateMarks(string rollNo, string paperCode, string series, string filename)
//        {
//            int maxMarks = 0;


//                        List<string> ansNoList = new List<string>();
//                        List<string> ansList = new List<string>();

//                using (SqlConnection connection = new SqlConnection(connectionString))
//                {
//                    connection.Open();
//                    SqlTransaction transaction = connection.BeginTransaction();

//                    try
//                    {

//                        string queryMaxMarks = "SELECT MAX(MAX_MARKS) AS MaxMarks FROM ANSKEY WHERE KEY_TYPE = @PaperCode AND SET_NO = @Series";
//                        using (SqlCommand command = new SqlCommand(queryMaxMarks, connection, transaction))
//                        {
//                            command.Parameters.AddWithValue("@PaperCode", paperCode);
//                            command.Parameters.AddWithValue("@Series", series);
//                            object result = command.ExecuteScalar();
//                            maxMarks = result != null && result != DBNull.Value ? Convert.ToInt32(result) : 0;
//                    }


//                        string queryAnswers = @"
//                   SELECT ANSKEY.ANS_NO, 
//       PaperInfo.ANS1, PaperInfo.ANS2, PaperInfo.ANS3, PaperInfo.ANS4, PaperInfo.ANS5, PaperInfo.ANS6, PaperInfo.ANS7, PaperInfo.ANS8, PaperInfo.ANS9, PaperInfo.ANS10,
//       PaperInfo.ANS11, PaperInfo.ANS12, PaperInfo.ANS13, PaperInfo.ANS14, PaperInfo.ANS15, PaperInfo.ANS16, PaperInfo.ANS17, PaperInfo.ANS18, PaperInfo.ANS19, PaperInfo.ANS20,
//       PaperInfo.ANS21, PaperInfo.ANS22, PaperInfo.ANS23, PaperInfo.ANS24, PaperInfo.ANS25, PaperInfo.ANS26, PaperInfo.ANS27, PaperInfo.ANS28, PaperInfo.ANS29, PaperInfo.ANS30,
//       PaperInfo.ANS31, PaperInfo.ANS32, PaperInfo.ANS33, PaperInfo.ANS34, PaperInfo.ANS35, PaperInfo.ANS36, PaperInfo.ANS37, PaperInfo.ANS38, PaperInfo.ANS39, PaperInfo.ANS40,
//       PaperInfo.ANS41, PaperInfo.ANS42, PaperInfo.ANS43, PaperInfo.ANS44, PaperInfo.ANS45, PaperInfo.ANS46, PaperInfo.ANS47, PaperInfo.ANS48, PaperInfo.ANS49, PaperInfo.ANS50,
//       PaperInfo.ANS51, PaperInfo.ANS52, PaperInfo.ANS53, PaperInfo.ANS54, PaperInfo.ANS55, PaperInfo.ANS56, PaperInfo.ANS57, PaperInfo.ANS58, PaperInfo.ANS59, PaperInfo.ANS60,
//       PaperInfo.ANS61, PaperInfo.ANS62, PaperInfo.ANS63, PaperInfo.ANS64, PaperInfo.ANS65, PaperInfo.ANS66, PaperInfo.ANS67, PaperInfo.ANS68, PaperInfo.ANS69, PaperInfo.ANS70,
//       PaperInfo.ANS71, PaperInfo.ANS72, PaperInfo.ANS73, PaperInfo.ANS74, PaperInfo.ANS75, PaperInfo.ANS76, PaperInfo.ANS77, PaperInfo.ANS78, PaperInfo.ANS79, PaperInfo.ANS80,
//       PaperInfo.ANS81, PaperInfo.ANS82, PaperInfo.ANS83, PaperInfo.ANS84, PaperInfo.ANS85, PaperInfo.ANS86, PaperInfo.ANS87, PaperInfo.ANS88, PaperInfo.ANS89, PaperInfo.ANS90,
//       PaperInfo.ANS91, PaperInfo.ANS92, PaperInfo.ANS93, PaperInfo.ANS94, PaperInfo.ANS95, PaperInfo.ANS96, PaperInfo.ANS97, PaperInfo.ANS98, PaperInfo.ANS99, PaperInfo.ANS100,
//       PaperInfo.ANS101, PaperInfo.ANS102, PaperInfo.ANS103, PaperInfo.ANS104, PaperInfo.ANS105, PaperInfo.ANS106, PaperInfo.ANS107, PaperInfo.ANS108, PaperInfo.ANS109, PaperInfo.ANS110,
//       PaperInfo.ANS111, PaperInfo.ANS112, PaperInfo.ANS113, PaperInfo.ANS114, PaperInfo.ANS115, PaperInfo.ANS116, PaperInfo.ANS117, PaperInfo.ANS118, PaperInfo.ANS119, PaperInfo.ANS120,
//       PaperInfo.ANS121, PaperInfo.ANS122, PaperInfo.ANS123, PaperInfo.ANS124, PaperInfo.ANS125, PaperInfo.ANS126, PaperInfo.ANS127, PaperInfo.ANS128, PaperInfo.ANS129, PaperInfo.ANS130,
//       PaperInfo.ANS131, PaperInfo.ANS132, PaperInfo.ANS133, PaperInfo.ANS134, PaperInfo.ANS135, PaperInfo.ANS136, PaperInfo.ANS137, PaperInfo.ANS138, PaperInfo.ANS139, PaperInfo.ANS140,
//       PaperInfo.ANS141, PaperInfo.ANS142, PaperInfo.ANS143, PaperInfo.ANS144, PaperInfo.ANS145, PaperInfo.ANS146, PaperInfo.ANS147, PaperInfo.ANS148, PaperInfo.ANS149, PaperInfo.ANS150
//FROM ANSKEY
//JOIN PaperInfo ON ANSKEY.KEY_TYPE = PaperInfo.[PAPER CODE]
//               AND ANSKEY.SET_NO = PaperInfo.[SERIES]
//               AND PaperInfo.[ROLL NO] = @RollNo";
//                    if (!string.IsNullOrEmpty(filename))
//                    {
//                        queryAnswers += " AND PaperInfo.filename = @filename";
//                    }

//                    queryAnswers += @"
//                WHERE ANSKEY.KEY_TYPE = @paperCode
//                  AND ANSKEY.SET_NO = @series
//                ORDER BY ANSKEY.Ques;";
//                    using (SqlCommand command = new SqlCommand(queryAnswers, connection, transaction))
//                        {
//                            command.Parameters.AddWithValue("@RollNo", rollNo);
//                            command.Parameters.AddWithValue("@PaperCode", paperCode);
//                            command.Parameters.AddWithValue("@Series", series);
//                        //command.Parameters.AddWithValue("@filename", filename);
//                        if (!string.IsNullOrEmpty(filename))
//                        {
//                            command.Parameters.AddWithValue("@filename", filename);
//                        }
//                        using (SqlDataReader reader = command.ExecuteReader())
//                            {
//                                while (reader.Read())
//                                {
//                                    ansNoList.Add(reader["ANS_NO"].ToString());
//                                    for (int i = 1; i <= 150; i++)
//                                    {
//                                        ansList.Add(reader["ANS" + i].ToString());
//                                    }
//                                }
//                            }
//                        }

//                        // Calculate marks
//                        int count = 0;
//                        for (int i = 0; i < ansNoList.Count; i++)
//                        {
//                            string[] correctAnswers = ansNoList[i].Split(',');
//                            string studentAnswer = ansList[i];
//                            if (correctAnswers.Any(ca => ca.Trim().Equals(studentAnswer, StringComparison.OrdinalIgnoreCase)))
//                            {
//                                count++;
//                            }
//                        }

//                        int getMarks = maxMarks * count;
//                    string updateQuery = @"
//                UPDATE PaperInfo 
//                SET MARKS = @Marks 
//                WHERE [ROLL NO] = @RollNo AND [PAPER CODE] = @PaperCode AND [SERIES] = @Series";

//                    if (!string.IsNullOrEmpty(filename))
//                    {
//                        updateQuery += " AND filename = @filename";
//                    }
//                    using (SqlCommand command = new SqlCommand(updateQuery, connection, transaction))
//                    {
//                        command.Parameters.AddWithValue("@Marks", getMarks);
//                        command.Parameters.AddWithValue("@RollNo", rollNo);
//                        command.Parameters.AddWithValue("@PaperCode", paperCode);
//                        command.Parameters.AddWithValue("@Series", series);
//                        if (!string.IsNullOrEmpty(filename))
//                        {
//                            command.Parameters.AddWithValue("@filename", filename);
//                        }

//                        command.ExecuteNonQuery();
//                    }


//                    transaction.Commit();
//                    }
//                    catch
//                    {
//                        transaction.Rollback();
//                        throw;
//                    }
//                }

//        }
 





git init
git add README.md
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/GitMohit-9621/CORRECTION_FORM.git
git push -u origin main