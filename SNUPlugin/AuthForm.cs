using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;

namespace SNUPlugin
{
    class AuthForm
    {
        //authorization form
        public static string getToken(string url, string msg)
        {
            Form myForm = new Form();
            Label lblDesc = new Label();
            TextBox txtInput = new TextBox();
            Button btnOK = new Button();
            Button btnCancel = new Button();
            bool boolCancel = true;
            Action actOK = delegate
            {
                if (!txtInput.Text.Trim().Equals(string.Empty))
                    boolCancel = false;
                myForm.Close();
            };
            myForm.Text = "추가 권한 요청...";
            myForm.ClientSize = new Size(380, 116);
            myForm.FormBorderStyle = FormBorderStyle.FixedDialog;
            myForm.StartPosition = FormStartPosition.CenterScreen;
            myForm.MaximizeBox = false;
            myForm.MinimizeBox = false;
            lblDesc.Text = msg;
            lblDesc.ForeColor = Color.Blue;
            lblDesc.Location = new Point(8, 8);
            lblDesc.Size = new Size(myForm.ClientSize.Width - 16, 36);
            lblDesc.Cursor = Cursors.Hand;
            lblDesc.Click += delegate
            {
                (new Process() { StartInfo = new ProcessStartInfo(url) { UseShellExecute = true } }).Start();
            };
            txtInput.Location = new Point(8, lblDesc.Location.Y + lblDesc.Size.Height + 8);
            txtInput.Size = new Size(myForm.ClientSize.Width - 16, 21);
            txtInput.PasswordChar = '*';
            txtInput.KeyUp += delegate (object sender, KeyEventArgs e)
            {
                if (e.KeyData == Keys.Return)
                    actOK.Invoke();
            };
            btnOK.Location = new Point(myForm.ClientSize.Width / 2 - 96 - 4, txtInput.Location.Y + txtInput.Size.Height + 8);
            btnOK.Size = new Size(96, 27);
            btnOK.Text = "확인";
            btnOK.Click += delegate { actOK.Invoke(); };
            btnCancel.Location = new Point(myForm.ClientSize.Width / 2 + 4, txtInput.Location.Y + txtInput.Size.Height + 8);
            btnCancel.Size = new Size(96, 27);
            btnCancel.Text = "취소";
            btnCancel.Click += delegate { myForm.Close(); };
            myForm.Controls.Add(lblDesc);
            myForm.Controls.Add(txtInput);
            myForm.Controls.Add(btnOK);
            myForm.Controls.Add(btnCancel);
            myForm.ShowDialog();
            if (boolCancel)
                return null;
            else
                return txtInput.Text;
        }
    }
}
