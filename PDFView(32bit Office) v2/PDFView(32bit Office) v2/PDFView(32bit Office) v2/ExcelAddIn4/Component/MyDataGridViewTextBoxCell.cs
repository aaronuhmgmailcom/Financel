using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelAddIn4.Component
{
    public partial class MyDataGridViewTextBoxCell : DataGridViewTextBoxCell
    {
        public MyDataGridViewTextBoxCell()
            : base() { }
        private DataGridViewTextBoxEditingControl dgvtbec;
        private DataGridViewColumn dgvc;
        private MyDataGridViewColumn mdgvc;
        public override void InitializeEditingControl(int rowIndex, object initialFormattedValue, DataGridViewCellStyle dataGridViewCellStyle)
        {
            base.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle);
            dgvtbec = DataGridView.EditingControl as DataGridViewTextBoxEditingControl;
            dgvc = this.OwningColumn;
            if (dgvc is MyDataGridViewColumn)
            {
                mdgvc = dgvc as MyDataGridViewColumn;
                dgvtbec.TextChanged += new EventHandler(dgvtbec_TextChanged);
            }
        }
        public void dgvtbec_TextChanged(object sender, EventArgs e)
        {
            mdgvc.DataGridViewColumnTextValue = dgvtbec.Text;
            EventCellChangeArgs ee = new EventCellChangeArgs(this.DataGridView.CurrentCell.RowIndex, this.DataGridView.CurrentCell.ColumnIndex, dgvtbec.Text);
            mdgvc.MyDataGridViewColumn_DataGridViewTextChanged(sender, ee);
        }
    }
    public class EventCellChangeArgs : EventArgs
    {
        public int rowIndex;
        public int columnIndex;
        public string value;
        public EventCellChangeArgs(int r, int c, string v)
        {
            rowIndex = r;
            columnIndex = c;
            value = v;
        }
    }
    public class MyDataGridViewColumn : DataGridViewColumn
    {
        public MyDataGridViewColumn()
            : base()
        {
            this.CellTemplate = new MyDataGridViewTextBoxCell();
        }
        public override DataGridViewCell CellTemplate
        {
            get
            {
                return base.CellTemplate;
            }
            set
            {
                if (value != null && !value.GetType().IsAssignableFrom(typeof(MyDataGridViewTextBoxCell)))
                {
                    throw new Exception("MyDataGridViewTextBoxCell");
                }
                base.CellTemplate = value;
            }
        }
        private string m_dataGridViewColumnTextValue = "";
        public string DataGridViewColumnTextValue
        {
            get
            {
                return m_dataGridViewColumnTextValue;
            }
            set
            {
                m_dataGridViewColumnTextValue = value;
            }
        }
        public void MyDataGridViewColumn_DataGridViewTextChanged(object sender, EventCellChangeArgs e)
        {
            if (DataGridViewTextChanged != null)
            {
                DataGridViewTextChanged(sender, e);
            }
        }
        public event EventCellChangeEvent DataGridViewTextChanged;
    }
    public delegate void EventCellChangeEvent(object sender, EventCellChangeArgs e);
}
