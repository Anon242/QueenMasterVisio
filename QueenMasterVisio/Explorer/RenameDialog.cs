using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QueenMasterVisio
{
    internal class RenameDialog
    {
        public static string ShowRenameDialog(string currentName)
        {
            using (var form = new Form())
            {
                form.Text = "Переименовать";
                form.Size = new Size(300, 120);
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.StartPosition = FormStartPosition.CenterParent;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                var textBox = new System.Windows.Forms.TextBox() { Text = currentName, Left = 20, Top = 20, Width = 240 };
                var buttonOk = new System.Windows.Forms.Button() { Text = "OK", Left = 130, Top = 50, Width = 60 };
                var buttonCancel = new System.Windows.Forms.Button() { Text = "Отмена", Left = 200, Top = 50, Width = 60 };
                var label = new System.Windows.Forms.Label() { Left = 20, Top = 3, Width = 240 };
                label.ForeColor = Color.Orange;

                buttonOk.DialogResult = DialogResult.OK;
                buttonCancel.DialogResult = DialogResult.Cancel;

                form.Controls.AddRange(new Control[] { textBox, buttonOk, buttonCancel, label });
                form.AcceptButton = buttonOk;
                form.CancelButton = buttonCancel;

                textBox.TextChanged += (object sender, EventArgs e) =>
                {
                    RenameDialogTextboxEvent(textBox, label, buttonOk);
                };
                form.Load += (object sender, EventArgs e) =>
                {
                    RenameDialogTextboxEvent(textBox, label, buttonOk);
                };

                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    return textBox.Text;
                }
            }

            return currentName;
        }
        private static void RenameDialogTextboxEvent(System.Windows.Forms.TextBox textBox, System.Windows.Forms.Label label, System.Windows.Forms.Button button)
        {
            // Тут логика проверки 
            if (textBox == null || string.IsNullOrEmpty(textBox.Text))
            {
                label.Text = "";
                button.Enabled = false;
                return;
            }
            label.ForeColor = Color.Orange;
            button.Enabled = true;

            Regex deviceReg = new Regex(@"^G\d");
            Regex lightReg = new Regex(@"^L\d");

            if (textBox.Text[0] == ' ' || textBox.Text[textBox.Text.Length - 1] == ' ')
            {
                label.ForeColor = Color.OrangeRed;
                label.Text = "Имя содержит в конце или в начале пробел";
                button.Enabled = false;
            }
            else if (deviceReg.IsMatch(textBox.Text))
            {
                label.Text = "Это устройство";
                int count = DevicesCountRegex(textBox.Text);
                if (count != 0)
                {
                    label.Text += " " + count + " шт.";
                }


            }
            else if (lightReg.IsMatch(textBox.Text))
            {
                label.Text = "Это cвет";
                int count = DevicesCountRegex(textBox.Text);
                if (count != 0)
                {
                    label.Text += " " + count + " шт.";
                }
            }
            else
            {
                label.ForeColor = Color.Gray;
                label.Text = "Другое";
            }

        }

        private static int DevicesCountRegex(string text)
        {
            Regex regex = new Regex(@"(G|L|g|l)(\d+)-(G|L|g|l)(\d+)");
            Match match = regex.Match(text);

            if (match.Success)
            {
                try
                {
                    int firstNumber = int.Parse(match.Groups[2].Value);
                    int secondNumber = int.Parse(match.Groups[4].Value);
                    return Math.Abs(firstNumber - secondNumber) + 1;
                }
                catch
                {
                    return 0;
                }
            }

            regex = new Regex(@"(G|L|g|l)(\d+)\.(\d+)-(G|L|g|l)?(\d+)\.(\d+)");
            match = regex.Match(text);

            if (match.Success)
            {
                try
                {
                    int firstNumber1 = int.Parse(match.Groups[2].Value);
                    int firstNumber2 = int.Parse(match.Groups[3].Value);
                    int secondNumber1 = int.Parse(match.Groups[5].Value);
                    int secondNumber2 = int.Parse(match.Groups[6].Value);

                    if (firstNumber1 - secondNumber1 != 0)
                        return 0;

                    return Math.Abs(firstNumber2 - secondNumber2) + 1;
                }
                catch
                {
                    return 0;
                }
            }
            return 0;
        }

    }
}
