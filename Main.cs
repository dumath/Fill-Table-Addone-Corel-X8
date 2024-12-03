using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Corel.Interop.VGCore;

namespace Fill_Table
{
    public partial class Main
    {
        // Врезка для себя. Пункт = 0.3528f;

        private int startNum;

        private double cellHeight;
        private double cellWidth;

        private DialogResult requirmentResponse;

        private void Startup()
        {
            if (app.ActiveDocument != null)
            {
                app.ActiveDocument.Unit = cdrUnit.cdrMillimeter;

                try
                {
                    UserDialog requirement = new UserDialog();
                    requirement.ShowDialog();

                    requirmentResponse = requirement.DialogResult;

                    if (requirement.DialogResult == DialogResult.OK)
                    {
                        startNum = requirement.startNumber;
                        cellHeight = requirement.height;
                        cellWidth = requirement.width;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Сначала нужно создать документ.", caption:"Уведомление");
            }
        }

        [CgsAddInMacro]
        public void Macro1()
        {
            if (requirmentResponse == DialogResult.OK)
            {
                if (app.ActiveDocument != null)
                {
                    if (app.ActiveLayer.Shapes.Count != 0)
                    {
                        DialogResult responseDeleting = MessageBox.Show("Документ не пустой. Очистить документ и запустить создание таблицы?", "Подтверждение", MessageBoxButtons.OKCancel);

                        if (responseDeleting == DialogResult.Cancel)
                        {
                            return;
                        }
                        else
                        {
                            string currentName = app.ActiveLayer.Name;

                            app.ActiveLayer.Delete();
                            app.ActivePage.CreateLayer(currentName);
                        }
                    }

                    fillTable();
                }
            }
        }

        private double remap(double width)
        {
            width /= cellWidth;
            width = Math.Truncate(width);

            return width *= cellWidth;
        }

        private void fillTable()
        {
            double remapWidth = remap(app.ActivePage.SizeWidth);

            Shape createdShape;

            double offsetX;
            double offsetY;

            for (double positionY = app.ActivePage.TopY; positionY > app.ActivePage.BottomY; positionY -= cellHeight)
            {
                for (double positionX = app.ActivePage.LeftX; positionX < remapWidth; positionX += cellWidth)
                {
                    offsetX = positionX + cellWidth;
                    offsetY = positionY - cellHeight;

                    createdShape = app.ActiveLayer.CreateParagraphText(positionX, positionY, offsetX, offsetY, startNum.ToString());

                    createdShape.Text.AlignProperties.Alignment = cdrAlignment.cdrCenterAlignment;
                    createdShape.Text.Frame.VerticalAlignment = cdrVerticalAlignment.cdrCenterJustify;

                    startNum++;
                }
            }
        }
    }
}
