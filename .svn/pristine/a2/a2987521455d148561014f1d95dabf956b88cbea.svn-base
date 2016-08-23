// Assignment implementation written by Conan Zhang, u0409453
// for CS3500 Assignment #6. November, 2014.
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SS;

namespace SpreadsheetGUI
{
    /// <summary>
    /// Main code for spreadsheet GUI operation.
    /// </summary>
    public partial class GUI : Form
    {
        //data structure behind GUI
        private Spreadsheet ss;

        //used to store previous cell for entry
        private int prevCol;
        private int prevRow;

        /// <summary>
        /// GUI constructor.
        /// </summary>
        public GUI()
        {
            InitializeComponent();

            ss = new Spreadsheet(s => true, s => s.ToUpper(), "ps6");
            prevCol = 0;
            prevRow = 0;

            //focus on contents
            this.ActiveControl = contentsBox;

            //add functionality for when user selects cell
            spreadsheetPanel.SelectionChanged += cellSelected;

            //set default selection
            DefaultSelection();
        }
        
        /// <summary>
        /// Triggered when File->New is clicked to open newwindow.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">a mouse click on File->New</param>
        private void newToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            //creates new spreadsheet GUI on same thread
            SpreadsheetApplicationContext.getAppContext().RunForm(new GUI());
        }

        /// <summary>
        /// Actions for when a cell is selected
        /// </summary>
        /// <param name="ss">spread sheet that selected cell is in</param>
        private void cellSelected(SpreadsheetPanel ssp)
        {
            //clear value display
            valueBox.Text = "";

            if (contentsBox.Text != "")
            {
                SetDisplayContentsOfCell();
            }

            UpdateCellDisplay();
        }

        /// <summary>
        /// Lets us intercept tab press.
        /// </summary>
        /// <param name="forward"></param>
        /// <returns>boolean if intercept worked</returns>
        protected override bool ProcessTabKey(bool forward)
        {
            //user moves and contents isn't empty so update
            if (contentsBox.Text != "")
            {
                //determine current selected cell
                spreadsheetPanel.GetSelection(out prevCol, out prevRow);

                SetDisplayContentsOfCell();

                MoveSelectionRight();
            }
            //no need to update
            else
            {
                MoveSelectionRight();
            }

            //intercept worked
            return true;
        }

        /// <summary>
        /// Lets us detect shortcut keyboard presses.
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="keyData">keyboard presses</param>
        /// <returns>boolean if intercept worked</returns>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            //user moves and contents isn't empty so update
            if (keyData == (Keys.Tab | Keys.Shift) && contentsBox.Text != "")
            {
                //determine current selected cell
                spreadsheetPanel.GetSelection(out prevCol, out prevRow);

                SetDisplayContentsOfCell();

                MoveSelectionLeft();
                return true;
            }
            //no need to update
            else if (keyData == (Keys.Tab | Keys.Shift))
            {
                MoveSelectionLeft();
                return true;
            }

            //if not intercepted call default
            return base.ProcessCmdKey(ref msg, keyData);
        }

        /// <summary>
        /// Waits for enter key to input contents of cell
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">user hits enter key</param>
        private void contentsBox_KeyDown_1(object sender, KeyEventArgs e)
        {
            //user moves and contents isn't empty so update
            if (contentsBox.Text != "")
            {
                if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Down)
                {
                    //determine current selected cell
                    spreadsheetPanel.GetSelection(out prevCol, out prevRow);

                    SetDisplayContentsOfCell();

                    MoveSelectionDown();
                }
                else if (e.KeyCode == Keys.Up)
                {
                    e.SuppressKeyPress = true;

                    //determine current selected cell
                    spreadsheetPanel.GetSelection(out prevCol, out prevRow);

                    SetDisplayContentsOfCell();

                    MoveSelectionUp();
                }
                else if (e.KeyCode == Keys.Right)
                {
                    e.SuppressKeyPress = true;

                    //determine current selected cell
                    spreadsheetPanel.GetSelection(out prevCol, out prevRow);

                    SetDisplayContentsOfCell();

                    MoveSelectionRight();
                }
                else if (e.KeyCode == Keys.Left)
                {
                    e.SuppressKeyPress = true;

                    //determine current selected cell
                    spreadsheetPanel.GetSelection(out prevCol, out prevRow);

                    SetDisplayContentsOfCell();

                    MoveSelectionLeft();
                }
            }
            //no need to update
            else
            {
                if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Down)
                {
                    e.SuppressKeyPress = true;
                    MoveSelectionDown();
                }
                else if (e.KeyCode == Keys.Up)
                {
                    e.SuppressKeyPress = true;
                    MoveSelectionUp();
                }
                else if (e.KeyCode == Keys.Right)
                {
                    e.SuppressKeyPress = true;
                    MoveSelectionRight();
                }
                else if (e.KeyCode == Keys.Left)
                {
                    e.SuppressKeyPress = true;
                    MoveSelectionLeft();
                }
            }
        }

        /// <summary>
        /// Creates save file dialog window.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">File->Save is clicked</param>
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveDialog();
        }

        /// <summary>
        /// Recieves file name from save file dialog.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">user hits ok</param>
        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            try
            {
                ss.Save(saveFileDialog.FileName);
            }
            catch
            {
                MessageBox.Show("Unable to save file.");
            }
        }

        /// <summary>
        /// Recieves and uses file name from open file dialog.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">user selects file</param>
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            try
            {
                //empty current spreadsheet display
                EmptySpreadsheet();

                //get data from file
                ss = new Spreadsheet(openFileDialog.FileName,s => true, s => s.ToUpper(), "ps6");

                //load cell data
                UpdateNonEmptyCells();

                //set selection to default
                DefaultSelection();
            }
            //unable to read file
            catch
            {
                MessageBox.Show("Unable to load specified file.");
            }
            
        }

        /// <summary>
        /// Creates open dialog window.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">File->Open is clicked</param>
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //no changes that could be overwritten
            if (ss.Changed != true)
            {
                OpenDialog();
            }
            else
            {
                DialogResult dialogResult = MessageBox.Show("You have unsaved changes in your spreadsheet. Would you like to save them?", "Spreadsheet", MessageBoxButtons.YesNoCancel);

                //user wants to save changes
                if (dialogResult == DialogResult.Yes)
                {
                    SaveDialog();
                }
                else if (dialogResult == DialogResult.No)
                {
                    OpenDialog();
                }
            }
        }

        /// <summary>
        /// User closes window.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">user hits close on window</param>
        private void GUI_FormClosing(object sender, FormClosingEventArgs e)
        {
            //could result in loss of unsaved data
            if (ss.Changed == true)
            {
                DialogResult dialogResult = MessageBox.Show("You have unsaved changes in your spreadsheet. Would you like to save them?", "Spreadsheet", MessageBoxButtons.YesNoCancel);

                //user wants to save changes
                if (dialogResult == DialogResult.Yes)
                {
                    SaveDialog();
                }
                //user wants to cancel closing
                else if (dialogResult == DialogResult.Cancel)
                {
                    e.Cancel = true;
                    return;
                }
            }
        }

        /// <summary>
        /// Triggered when File->Close is clicked to close current window.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">a mouse click on File->Close</param>
        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Opens help website for user.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">user clicks on Website->Help</param>
        private void websiteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"..\..\..\Resources\HelpPage.html");
        }

        /// <summary>
        /// Used to move selection down when necessary.
        /// </summary>
        private void MoveSelectionDown()
        {
            if (prevRow != spreadsheetPanel.Height)
            {
                //selection goes down
                spreadsheetPanel.SetSelection(prevCol, prevRow + 1);

                UpdateCellDisplay();
            }
        }

        /// <summary>
        /// Used to move selection up when necessary.
        /// </summary>
        private void MoveSelectionUp()
        {
            if (prevRow != 0)
            {
                //selection goes down
                spreadsheetPanel.SetSelection(prevCol, prevRow - 1);

                UpdateCellDisplay();
            }
        }

        /// <summary>
        /// Used to move selection right when necessary.
        /// </summary>
        private void MoveSelectionRight()
        {
            if (prevCol != spreadsheetPanel.Width)
            {
                //selection goes down
                spreadsheetPanel.SetSelection(prevCol + 1, prevRow);

                UpdateCellDisplay();
            }
        }

        /// <summary>
        /// Used to move selection left when necessary.
        /// </summary>
        private void MoveSelectionLeft()
        {
            if (prevCol != 0)
            {
                //selection goes down
                spreadsheetPanel.SetSelection(prevCol - 1, prevRow);

                UpdateCellDisplay();
            }
        }

        /// <summary>
        /// Updates display of cell name, value, and contents.
        /// </summary>
        private void UpdateCellDisplay()
        {
            //determine current selected cell
            spreadsheetPanel.GetSelection(out prevCol, out prevRow);

            //display current cell name
            nameBox.Text = ColRowName(prevCol, prevRow);

            //display current cell value
            valueBox.Text = ss.GetCellValue(nameBox.Text).ToString();

            //display current cell contents
            contentsBox.Text = ss.GetCellContents(nameBox.Text).ToString();
            //cursor starts at end of line
            contentsBox.Select(contentsBox.TextLength, 0); 
        }

        /// <summary>
        /// Sets GUI and data of cell.
        /// </summary>
        private void SetDisplayContentsOfCell()
        {
            //setting could cause circular exception
            try
            {
                //set contents of cell
                foreach (string name in ss.SetContentsOfCell(ColRowName(prevCol, prevRow), contentsBox.Text))
                {
                    SetCellData(name, ss.GetCellValue(name).ToString());
                }

                string value = ss.GetCellValue(ColRowName(prevCol, prevRow)).ToString();
                //set display value to contents
                spreadsheetPanel.SetValue(prevCol, prevRow, value);

                //set display value
                valueBox.Text = value;
            }
            //circular exception
            catch(CircularException)
            {
                MessageBox.Show("A circular dependency occurs when this action is performed!");
            }
            //formula format exception
            catch
            {
                MessageBox.Show("Formula is incorrectly formatted. Please check syntax for any errors.");
            }
        }

        /// <summary>
        /// Sets cell selection to default.
        /// </summary>
        private void DefaultSelection()
        {
            spreadsheetPanel.SetSelection(0, 0);

            //display
            nameBox.Text = "A1";
            valueBox.Text = ss.GetCellValue("A1").ToString();
            contentsBox.Text = ss.GetCellContents("A1").ToString();

            spreadsheetPanel.GetSelection(out prevCol, out prevRow);
            //cursor starts at end of line
            contentsBox.Select(contentsBox.TextLength, 0);
        }

        /// <summary>
        /// Sets all spreadsheet values to an empty string.
        /// </summary>
        private void EmptySpreadsheet()
        {
            foreach (string name in ss.GetNamesOfAllNonemptyCells())
            {
                SetCellData(name, "");
            }
        }
        /// <summary>
        /// Updates all nonempty cells in spreadsheet.
        /// </summary>
        private void UpdateNonEmptyCells()
        {
            foreach (string name in ss.GetNamesOfAllNonemptyCells())
            {
                SetCellData(name, ss.GetCellValue(name).ToString());
            }
        }

        /// <summary>
        /// Set data of cell given by name.
        /// </summary>
        /// <param name="name">name of cell</param>
        /// <param name="data">data to set cell to</param>
        private void SetCellData(string name, string data)
        {
            //get name of cell
            int col = (int)name[0] % 32 - 1;
            int row;
            bool intParse = Int32.TryParse(name.Remove(0, 1), out row);

            //set display value
            spreadsheetPanel.SetValue(col, row - 1, data);
        }

        /// <summary>
        /// Gets name of cell based on row and column numbers.
        /// </summary>
        /// <param name="col">column number of cell</param>
        /// <param name="row">row number of cell</param>
        /// <returns>name of cell</returns>
        private string ColRowName(int col, int row)
        {
            return IntToString(col) + (row + 1).ToString();
        }

        /// <summary>
        /// Converts integer to corresponding letter for display purposes
        /// </summary>
        /// <param name="num">integer to convert to corresponding string</param>
        /// <returns>string corresponding to integer</returns>
        private String IntToString(int num)
        {
            Char c = (Char)(65 + num);
            return c.ToString();
        }

        /// <summary>
        /// Used to bring up save file dialog.
        /// </summary>
        private void SaveDialog()
        {
            saveFileDialog.FileName = "ps6.sprd";
            saveFileDialog.Filter = "Spreadsheet files (*.sprd)|*.sprd|All files (*.*)|*.*";
            saveFileDialog.ShowDialog();
        }

        /// <summary>
        /// Used to bring up open file dialog.
        /// </summary>
        private void OpenDialog()
        {
            openFileDialog.FileName = "";
            openFileDialog.Filter = "Spreadsheet files (*.sprd)|*.sprd|All files (*.*)|*.*";
            openFileDialog.ShowDialog();
        }
    }
}
