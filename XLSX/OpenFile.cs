using System.Windows.Forms;

namespace control
{
    class OpenFile
    {
        public void openFile()
        {
          OpenFileDialog openFile = new OpenFileDialog();
          openFile.Filter = "Формат xlsx(*.xlsx)|*.xlsx|Все файлы(*.*)|*.*";
          openFile.Title = "Выберете файл";

            if (openFile.ShowDialog() == DialogResult.OK)
            {
             Fields.FileName = openFile.FileName;
            }

          FileSelected fileSelected = new FileSelected();
          fileSelected.xlsxSelected();
        }
    }
}
