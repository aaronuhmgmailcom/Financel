using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using ExcelAddIn4;

namespace ExcelAddIn3
{
    public partial class Form1 : Form
    {
        internal static Finance_Tools ft
        {
            get { return new Finance_Tools(); }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_CellMouseDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dataGridView1[e.ColumnIndex, e.RowIndex].Value != null)
                {
                    if (dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString().Contains("$"))
                    {
                        string KeyWord = dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString().Replace("$", "");
                        dataGridView1[e.ColumnIndex, e.RowIndex].Value = "";
                        dataGridView1[e.ColumnIndex, e.RowIndex].Value = KeyWord;
                    }
                    else
                    {
                        string KeyWord = dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString();
                        string res = Regex.Replace(KeyWord, @"(\d+)|(\s+) ", " $1 $2 ", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                        KeyWord = "$" + res.Trim().Replace(" ", "$");
                        dataGridView1[e.ColumnIndex, e.RowIndex].Value = "";
                        dataGridView1[e.ColumnIndex, e.RowIndex].Value = KeyWord;
                    }
                    dataGridView1.EndEdit();
                }
            }
            catch
            {
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            System.Drawing.Rectangle rectangle = new System.Drawing.Rectangle(e.RowBounds.Location.X,
                e.RowBounds.Location.Y,
                dataGridView1.RowHeadersWidth - 4,
                e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                dataGridView1.RowHeadersDefaultCellStyle.Font,
                rectangle,
                dataGridView1.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindData()
        {
            DataTable dt = ft.GetLineDetailDataFromDB();
            this.dataGridView1.DataSource = dt;
            
            if (dt.Rows.Count > 0)
            {
                //this.txtLineIndicator.Text = dt.Rows[0]["LineIndicator"].ToString();
                //this.txtStartCell.Text = dt.Rows[0]["StartinginCell"].ToString();
            }
        }
        public Form1()
        {
            try
            {
                InitializeComponent();
                DatagridViewCheckBoxCell dgvcCheckBox = new DatagridViewCheckBoxCell();
                dgvcCheckBox.OnCheckBoxClicked += new CheckBoxClickedHandler(HeaderCellChecked);//注册列头的checkbox选中事件

                //dataGridView1 = ft.IniGrd();

                dataGridView1.AllowUserToAddRows = true;
                dataGridView1.CellDoubleClick += new DataGridViewCellEventHandler(dgv_CellMouseDoubleClick);
                //dataGridView1.CellClick += new DataGridViewCellEventHandler(dgv_CellMouseClick);
                dataGridView1.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dgv_RowPostPaint);
                //dataGridView1.colu
                //BindData();

                //dataGridView1.Columns.Add("Checked", "Line Checked");
                //dataGridView1.Columns.Add("Checked2", "Line Checked2");
                //dataGridView1.Columns.Add("Checked3", "Line Checked3");
                dataGridView1.Rows.Add("1");
                dataGridView1.Rows.Add("2");
                dataGridView1.Rows.Add("3");
                DataGridViewComboBoxCell combox = new  DataGridViewComboBoxCell();
                //combox.Items.Clear();
                combox.Items.Add("boy");
                combox.Items.Add("girl");

                dataGridView1.Rows[0].Cells[0] = combox;

                //dataGridView1.Columns.Add("Checked", "Line Checked");
                //dataGridView1.Columns["Checked"].DataPropertyName = "Checked";

                //dataGridView1.Columns["Checked"].CellType = DataGridViewCheckBoxCell;//设置checkbox列列头cell为我们画的那个checkbox列头
                //DataGridViewCheckBoxCell dgvcCheckBox2 = new DataGridViewCheckBoxCell();
                
                //dataGridView1.Rows[0].Cells.Add(dgvcCheckBox2);

              
                //dgvLD.LostFocus += new EventHandler(dgvLD_LostFocus);
                //dataGridView1.NotifyCurrentCellDirty(false);
                //dataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke;
                //this.panel1.Controls.Add(dataGridView1);
                //dataGridView1.KeyDown += new KeyEventHandler(dgvLD_KeyDown);
            }
            catch (Exception ex)
            {

            }
        }
        public void HeaderCellChecked(bool state)
        {
            //代码
        }
        public void HeaderCellChecked(object sender, DataGridViewCheckBoxCellEventArgs e)
        {
            //代码
        }
    }
    public class DatagridViewCheckBoxCell : DataGridViewCheckBoxCell
    {
        Point checkBoxLocation;
        Size checkBoxSize;
        bool _checked = false;
        Point _cellLocation = new Point();
        System.Windows.Forms.VisualStyles.CheckBoxState _cbState =
            System.Windows.Forms.VisualStyles.CheckBoxState.UncheckedNormal;
        public event CheckBoxClickedHandler OnCheckBoxClicked;

        public DatagridViewCheckBoxCell()
        {
        }

        public bool Checked
        {
            get { return _checked; }
            set
            {
                _checked = value;
                //OnCheckBoxClicked(_checked);
                this.DataGridView.InvalidateCell(this);
            }
        }

        protected override void Paint(System.Drawing.Graphics graphics,
            System.Drawing.Rectangle clipBounds,
            System.Drawing.Rectangle cellBounds,
            int rowIndex,
            DataGridViewElementStates dataGridViewElementState,
            object value,
            object formattedValue,
            string errorText,
            DataGridViewCellStyle cellStyle,
            DataGridViewAdvancedBorderStyle advancedBorderStyle,
            DataGridViewPaintParts paintParts)
        {
            base.Paint(graphics, clipBounds, cellBounds, rowIndex,
                dataGridViewElementState, value,
                "", errorText, cellStyle,
                advancedBorderStyle, paintParts);
            Point p = new Point();
            Size s = CheckBoxRenderer.GetGlyphSize(graphics,
            System.Windows.Forms.VisualStyles.CheckBoxState.UncheckedNormal);
            p.X = cellBounds.Location.X +
                (cellBounds.Width / 2) - (s.Width / 2);
            p.Y = cellBounds.Location.Y +
                (cellBounds.Height / 2) - (s.Height / 2);
            _cellLocation = cellBounds.Location;
            checkBoxLocation = p;
            checkBoxSize = s;
            if (_checked)
                _cbState = System.Windows.Forms.VisualStyles.
                    CheckBoxState.CheckedNormal;
            else
                _cbState = System.Windows.Forms.VisualStyles.
                    CheckBoxState.UncheckedNormal;
            CheckBoxRenderer.DrawCheckBox
            (graphics, checkBoxLocation, _cbState);
        }

        protected override void OnMouseClick(DataGridViewCellMouseEventArgs e)
        {
            Point p = new Point(e.X + _cellLocation.X, e.Y + _cellLocation.Y);
            if (p.X >= checkBoxLocation.X && p.X <=
                checkBoxLocation.X + checkBoxSize.Width
            && p.Y >= checkBoxLocation.Y && p.Y <=
                checkBoxLocation.Y + checkBoxSize.Height)
            {
                _checked = !_checked;
                if (OnCheckBoxClicked != null)
                {
                    OnCheckBoxClicked(_checked);
                    this.DataGridView.InvalidateCell(this);
                }
            }
            base.OnMouseClick(e);
        }
    }
    public delegate void CheckBoxClickedHandler(bool State);
    public class DataGridViewCheckBoxCellEventArgs : EventArgs
    {
        bool _bChecked;
        public DataGridViewCheckBoxCellEventArgs(bool bChecked)
        {
            _bChecked = bChecked;
        }
        public bool Checked
        {
            get { return _bChecked; }
        }
    }




}
