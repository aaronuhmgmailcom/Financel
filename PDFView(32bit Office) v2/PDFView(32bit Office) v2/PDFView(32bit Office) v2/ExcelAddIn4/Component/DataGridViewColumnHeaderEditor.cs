using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace Utility.DataComponent
{
    /// <summary>
    /// 
    /// </summary>
    public partial class DataGridViewColumnHeaderEditor : Component, ISupportInitialize
    {
        public DataGridViewColumnHeaderEditor()
        {
            InitializeComponent();
        }
        public DataGridViewColumnHeaderEditor(IContainer container)
        {
            container.Add(this);
            InitializeComponent();
        }
        #region
        /// <summary>
        /// 
        /// </summary>
        public event ColumnHeaderEditEventHandler BeginEdit;
        /// <summary>
        /// 
        /// </summary>
        public event ColumnHeaderEditEventHandler EndingEdit;
        /// <summary>
        /// ­
        /// </summary>
        public event ColumnHeaderEditEventHandler EndEdit;
        private int m_SelectedColumnIndex = -1;
        private int m_ScrollValue = 0;
        private SortedList<int, DataGridViewColumn> m_SortedColumnList = new SortedList<int, DataGridViewColumn>();
        private DataGridView m_TargetControl = null;
        [Description("")]
        public DataGridView TargetControl
        {
            get { return m_TargetControl; }
            set { m_TargetControl = value; }
        }
        private bool m_EnableEdit = true;
        [Description(""), DefaultValue(true)]
        public bool EnableEdit
        {
            get { return m_EnableEdit; }
            set { m_EnableEdit = value; }
        }
        #endregion
        #region ISupportInitialize
        public void BeginInit()
        {
        }
        public void EndInit()
        {
            if (m_TargetControl != null)
            {
                this.m_TargetControl.Parent.Controls.Add(this.rtbTitle);
                this.rtbTitle.BringToFront();
                this.ReloadSortedColumnList();
                m_TargetControl.ColumnHeaderMouseDoubleClick += new DataGridViewCellMouseEventHandler(TargetControl_ColumnHeaderMouseDoubleClick);
                m_TargetControl.ColumnDisplayIndexChanged += new DataGridViewColumnEventHandler(TargetControl_ColumnDisplayIndexChanged);
                m_TargetControl.ColumnRemoved += new DataGridViewColumnEventHandler(TargetControl_ColumnRemoved);
                m_TargetControl.ColumnWidthChanged += new DataGridViewColumnEventHandler(TargetControl_ColumnWidthChanged);
                m_TargetControl.ColumnAdded += new DataGridViewColumnEventHandler(TargetControl_ColumnAdded);
                m_TargetControl.Scroll += new ScrollEventHandler(TargetControl_Scroll);
            }
        }
        #endregion ISupportInitialize
        #region
        void TargetControl_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
                this.m_ScrollValue = e.NewValue;
            if (this.rtbTitle.Visible)
                this.ShowHeaderEdit();
        }
        void TargetControl_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            //throw new Exception("The method or operation is not implemented.");
        }
        void TargetControl_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            this.ReloadSortedColumnList();
        }
        void TargetControl_ColumnRemoved(object sender, DataGridViewColumnEventArgs e)
        {
            this.ReloadSortedColumnList();
        }
        void TargetControl_ColumnDisplayIndexChanged(object sender, DataGridViewColumnEventArgs e)
        {
            this.ReloadSortedColumnList();
        }
        void TargetControl_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            this.m_SelectedColumnIndex = this.m_TargetControl.Columns[e.ColumnIndex].DisplayIndex;
            if (this.m_EnableEdit)
                this.ShowHeaderEdit();
        }
        #endregion
        #region
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rtbTitle_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Enter:
                    this.m_TargetControl.Focus();
                    e.Handled = true;
                    break;
                case Keys.Right:
                    if (this.rtbTitle.SelectionStart >= this.rtbTitle.Text.Length)
                    {
                        if (this.m_SelectedColumnIndex < this.m_TargetControl.Columns.Count - 1)
                        {
                            e.Handled = true;
                            this.m_TargetControl.Focus();
                            this.m_SelectedColumnIndex = this.GetNextVisibleColumnIndex(this.m_SelectedColumnIndex);
                            this.ShowHeaderEdit();
                        }
                    }
                    break;
                case Keys.Left:
                    if (this.rtbTitle.SelectionStart == 0)
                    {
                        if (this.m_SelectedColumnIndex > 0)
                        {
                            e.Handled = true;
                            this.m_TargetControl.Focus();
                            this.m_SelectedColumnIndex = this.GetPreVisibleColumnIndex(this.m_SelectedColumnIndex);
                            this.ShowHeaderEdit();
                        }
                    }
                    break;
                default:
                    break;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rtbTitle_Leave(object sender, EventArgs e)
        {
            DataGridViewColumn myColumn = this.m_SortedColumnList[this.m_SelectedColumnIndex];
            ColumnHeaderEditEventArgs myArgs = new ColumnHeaderEditEventArgs(myColumn, this.rtbTitle.Text.Trim());
            if (this.EndingEdit != null)
            {
                this.EndingEdit(this, myArgs);
                if (myArgs.Cancel)
                {
                    this.rtbTitle.Focus();
                    return;
                }
            }
            this.rtbTitle.Visible = false;
            if (this.rtbTitle.Text.Trim().Length > 0)
            {
                if (myColumn.HeaderText != this.rtbTitle.Text.Trim())
                {
                    myColumn.HeaderText = this.rtbTitle.Text.Trim();
                }
            }
            if (this.EndEdit != null)
                this.EndEdit(this, myArgs);
        }
        #endregion
        #region
        /// <summary>
        /// 
        /// </summary>
        private void ReloadSortedColumnList()
        {
            this.m_SortedColumnList.Clear();
            foreach (DataGridViewColumn column in this.m_TargetControl.Columns)
            {
                this.m_SortedColumnList.Add(column.DisplayIndex, column);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ColumnIndex"></param>
        /// <returns></returns>
        private int GetColumnRelativeLeft(int ColumnIndex)
        {
            int intLeft = 0;
            DataGridViewColumn Column = null;

            for (int intIndex = 0; intIndex < ColumnIndex; intIndex++)
            {
                if (this.m_SortedColumnList.ContainsKey(intIndex))
                {
                    Column = this.m_SortedColumnList[intIndex];
                    if (Column.Visible)
                        intLeft += Column.Width + Column.DividerWidth;
                }
            }
            return intLeft;
        }
        /// <summary>
        /// 
        /// </summary>
        private void ShowHeaderEdit()
        {
            if (this.BeginEdit != null)
            {
                ColumnHeaderEditEventArgs myArgs = new ColumnHeaderEditEventArgs(this.m_SortedColumnList[this.m_SelectedColumnIndex], "");
                BeginEdit(this, myArgs);
                if (myArgs.Cancel)
                    return;
            }
            int intColumnRelativeLeft = 0;
            int intFirstColumnLeft = (this.m_TargetControl.RowHeadersVisible ? this.m_TargetControl.RowHeadersWidth + 1 : 1);
            int intTargetX = this.m_TargetControl.Location.X, intTargetY = this.m_TargetControl.Location.Y, intTargetWidth = this.m_TargetControl.Width;

            intColumnRelativeLeft = GetColumnRelativeLeft(this.m_SelectedColumnIndex);

            if (intColumnRelativeLeft < this.m_ScrollValue)
            {
                this.rtbTitle.Location = new Point(intTargetX + intFirstColumnLeft, intTargetY + 1);
                if (intColumnRelativeLeft + this.m_SortedColumnList[this.m_SelectedColumnIndex].Width > this.m_ScrollValue)
                    this.rtbTitle.Width = intColumnRelativeLeft + this.m_SortedColumnList[this.m_SelectedColumnIndex].Width - this.m_ScrollValue;
                else
                    this.rtbTitle.Width = 0;
            }
            else
            {
                this.rtbTitle.Location = new Point(intColumnRelativeLeft + intTargetX - this.m_ScrollValue + intFirstColumnLeft, intTargetY + 1);

                if (this.rtbTitle.Location.X + this.rtbTitle.Width > intTargetX + intTargetWidth)
                {
                    int intWidth = intTargetX + intTargetWidth - this.rtbTitle.Location.X;
                    this.rtbTitle.Width = (intWidth >= 0 ? intWidth : 0);
                }
                else
                    this.rtbTitle.Width = this.m_SortedColumnList[this.m_SelectedColumnIndex].Width;
            }
            this.rtbTitle.Height = this.m_TargetControl.ColumnHeadersHeight - 1;
            this.rtbTitle.Text = this.m_SortedColumnList[this.m_SelectedColumnIndex].HeaderText;
            this.rtbTitle.SelectAll();
            this.rtbTitle.Visible = true;
            this.rtbTitle.Focus();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="CurrentIndex"></param>
        /// <returns></returns>
        private int GetPreVisibleColumnIndex(int CurrentIndex)
        {
            int intPreIndex = 0;

            for (int intIndex = CurrentIndex - 1; intIndex >= 0; intIndex--)
            {
                if (this.m_SortedColumnList.ContainsKey(intIndex) && this.m_SortedColumnList[intIndex].Visible)
                {
                    intPreIndex = intIndex;
                    break;
                }
            }
            return intPreIndex;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="CurrentIndex"></param>
        /// <returns></returns>
        private int GetNextVisibleColumnIndex(int CurrentIndex)
        {
            int intNextIndex = CurrentIndex;
            for (int intIndex = CurrentIndex + 1; intIndex <= this.m_SortedColumnList.Keys[this.m_SortedColumnList.Count - 1]; intIndex++)
            {
                if (this.m_SortedColumnList.ContainsKey(intIndex) && this.m_SortedColumnList[intIndex].Visible)
                {
                    intNextIndex = intIndex;
                    break;
                }
            }
            return intNextIndex;
        }
        #endregion
    }//class DataGridViewColumnHeaderEditor
    public delegate void ColumnHeaderEditEventHandler(object sender, ColumnHeaderEditEventArgs e);
    public class ColumnHeaderEditEventArgs : EventArgs
    {
        private bool m_Cancel = false;
        /// <summary>
        /// ­
        /// </summary>
        public bool Cancel
        {
            get { return m_Cancel; }
            set { m_Cancel = value; }
        }

        private string m_NewHeaderText = "";
        /// <summary>
        /// 
        /// </summary>
        public string NewHeaderText
        {
            get { return m_NewHeaderText; }
            set
            {
                if (!(string.IsNullOrEmpty(value) || value.Trim().Length == 0))
                    m_NewHeaderText = value;
            }
        }
        private DataGridViewColumn m_Column = null;
        /// <summary>
        /// 
        /// </summary>
        public DataGridViewColumn Column
        {
            get { return m_Column; }
        }
        public ColumnHeaderEditEventArgs(DataGridViewColumn Column, string NewHeaderText)
        {
            if (Column == null)
                throw new ArgumentNullException("Column", "");
            this.m_Column = Column;
            if (string.IsNullOrEmpty(NewHeaderText) || NewHeaderText.Trim().Length == 0)
                NewHeaderText = Column.HeaderText;
            this.m_NewHeaderText = NewHeaderText.Trim();
        }
    }
}
