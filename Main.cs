using System;
using System.Linq;
using System.Windows.Forms;
using Corel.Interop.VGCore;

namespace Create_Editable_Cells
{
    public partial class Main
    {
        public const double K_SIDES = 2d; // Коэффициент граней.
        public const double K_PRVIEW_NxN = 3d; // Коэффициент Column x Row.

        private Document activeDocument; // Пока не используется. Привязка документа, на котором был вызван плагин.
        private Page activePage; // Пока не используется. Тоже самое.

        private Page previewPage;

        #region Требования пользователя.

        #region Число первой ячейки.
        private int startNumber;
        public string RefreshStartNumber(string text, bool isPreviw)
        {
            text = string.Concat(text.Where(x => char.IsDigit(x)));

            bool result = false;

            if (int.TryParse(text, out startNumber))
            {
                result = isAllowedCreateMap();

                if (result && isPreviw)
                    CreatePreviewMap();
            }

            return text;
        }
        #endregion

        #region Ширина ячейки.
        private double cellWidth = 0d;
        public bool RefreshCellWidth(string text, bool isPreview)
        {
            bool result = false;

            if (double.TryParse(text, out cellWidth))
            {
                result = isAllowedCreateMap();

                if (result && isPreview)
                    CreatePreviewMap();
            }

            return result;
        }
        #endregion

        #region Высота ячейки.
        private double cellHeight = 0d;
        public bool RefreshCellHeight(string text, bool isPreview)
        {
            bool result = false;

            if (double.TryParse(text, out cellHeight))
            {
                result = isAllowedCreateMap();

                if (result && isPreview)
                    CreatePreviewMap();
            }

            return result;
        }
        #endregion

        #region Отступ от края ячейки.
        private double margin = 0d;
        public bool RefreshMargin(decimal value, bool isPreview)
        {
            bool result = result = isAllowedCreateMap();

            margin = Convert.ToDouble(value);

            if (result && isPreview)
                CreatePreviewMap();

            return result;
        }
        #endregion

        #region Ширина абриса.
        private double outline = 0d;
        public bool RefreshOutline(string text, bool isPreview)
        {
            bool result = false;

            if (double.TryParse(text, out outline))
            {
                result = isAllowedCreateMap();

                if (result && isPreview)
                    CreatePreviewMap();
            }

            return result;
        }
        #endregion

        #endregion

        private void Startup() { if (app.ActiveDocument != null) app.ActiveDocument.Unit = cdrUnit.cdrMillimeter; }

        [CgsAddInMacro]
        public void Run()
        {
            if (app.ActiveDocument != null)
            {
                if (app.ActivePage.Shapes.Count > 0)
                {
                    // Пользователь должен сам решить, что удалить.
                    MessageBox.Show("Активный документ содержит элементы.", caption: "Уведомление");
                }
                else
                {
                    Preference userPreference = new Preference(this);
                    userPreference.ShowDialog();
                }
            }
            else
            {
                // Пользователь должен сам решить, какой формат листа нужно заполнять.
                MessageBox.Show("Сначала нужно создать документ.", caption: "Уведомление");
            }
        }

        public void CreateMap()
        {
            double yStartPosition = app.ActivePage.TopY - (margin + outline);
            double xStartPosition = app.ActivePage.LeftX + (margin + outline);

            double yOffset = cellHeight + (margin + outline) * K_SIDES;
            double xOffset = cellWidth + (margin + outline) * K_SIDES;

            double yEndPosition;
            double xEndPosition;

            remap(app.ActivePage, out xEndPosition, out yEndPosition, xOffset, yOffset);

            yEndPosition = app.ActivePage.SizeHeight - yEndPosition;

            Shape rect;
            Shape text;

            for (double yPos = yStartPosition; yPos > yEndPosition; yPos -= yOffset)
            {
                for (double xPos = xStartPosition; xPos < xEndPosition; xPos += xOffset)
                {
                    rect = app.ActiveLayer.CreateRectangle(xPos, yPos, xPos + cellWidth, yPos - cellHeight);

                    if(outline > 0d)
                    {
                        rect.Outline.SetProperties(Width: outline);
                        rect.Outline.Justification = cdrOutlineJustification.cdrOutlineJustificationOutside;
                    }
                    else
                    {
                        rect.Outline.SetNoOutline();
                    }

                    text = app.ActiveLayer.CreateParagraphText(rect.LeftX, rect.TopY, rect.RightX, rect.BottomY, Text: startNumber.ToString());
                    text.Text.AlignProperties.Alignment = cdrAlignment.cdrCenterAlignment;
                    text.Text.Frame.VerticalAlignment = cdrVerticalAlignment.cdrCenterJustify;

                    rect.PlaceTextInside(text);

                    startNumber++;
                }
            }
        }

        public void CreatePreviewMap()
        {
            if (isAllowedCreateMap())
            {
                int previewStartNumber = startNumber;

                double pageWidth = (cellWidth + (margin + outline) * K_SIDES) * K_PRVIEW_NxN;
                double pageHeight = (cellHeight + (margin + outline) * K_SIDES) * K_PRVIEW_NxN;

                сreatePreviewPage(pageWidth, pageHeight);

                double yStartPosition = app.ActivePage.TopY - (margin + outline);
                double xStartPosition = app.ActivePage.LeftX + (margin + outline);

                double yEndPosition = app.ActivePage.BottomY;
                double xEndPosition = app.ActivePage.RightX;

                double yOffset = cellHeight + (margin + outline) * K_SIDES;
                double xOffset = cellWidth + (margin + outline) * K_SIDES;

                Shape rect;
                Shape text;

                for (double yPos = yStartPosition; yPos > yEndPosition; yPos -= yOffset)
                {
                    for (double xPos = xStartPosition; xPos < xEndPosition; xPos += xOffset)
                    {
                        rect = app.ActiveLayer.CreateRectangle(xPos, yPos, xPos + cellWidth, yPos - cellHeight);

                        if (outline > 0d)
                        {
                            rect.Outline.SetProperties(Width: outline);
                            rect.Outline.Justification = cdrOutlineJustification.cdrOutlineJustificationOutside;
                        }
                        else
                        {
                            rect.Outline.SetNoOutline();
                        }

                        text = app.ActiveLayer.CreateParagraphText(rect.LeftX, rect.TopY, rect.RightX, rect.BottomY, Text: previewStartNumber.ToString());
                        text.Text.AlignProperties.Alignment = cdrAlignment.cdrCenterAlignment;
                        text.Text.Frame.VerticalAlignment = cdrVerticalAlignment.cdrCenterJustify;

                        rect.PlaceTextInside(text);

                        previewStartNumber++;
                    }
                }
            }
        }

        /// <summary>
        /// Удаляет страницу превью из документа.
        /// </summary>
        public void RemovePreviewPage()
        {
            if (previewPage != null)
            {
                app.ActiveDocument.Pages.First.Activate();

                previewPage.Delete();
                previewPage = null;
            }
        }

        #region Дополнительные методы Макроса. Внутренние.
        /// <summary>
        /// Можно ли строить карту ячеек. Смотрит введеные значения пользователя на двух основных свойствах.
        /// </summary>
        /// <returns> Резулатат проверки. </returns>
        private bool isAllowedCreateMap()
        {
            bool result = false;

            if (cellWidth > 0d && cellHeight > 0d)
                result = true;

            return result;
        }

        /// <summary>
        /// Создает отдельную превью - страницу в текущем документе.
        /// </summary>
        /// <param name="pageWidth"> Ширина страницы </param>
        /// <param name="pageHeight"> Высота страницы </param>
        private void сreatePreviewPage(double pageWidth, double pageHeight)
        {
            if (previewPage == null)
            {
                previewPage = app.ActiveDocument.AddPagesEx(1, pageWidth, pageHeight);
                previewPage.Activate();
            }
            else
            {
                previewPage.SetSize(pageWidth, pageHeight);
                previewPage.Activate();

                if (previewPage.ActiveLayer.Shapes.Count > 0)
                {
                    previewPage.ActiveLayer.Delete();
                    previewPage.CreateLayer("Preview_Layer");
                    previewPage.ActiveLayer.Activate();
                }
            }
        }

        /// <summary>
        /// Получает размер заполненной страницы ячейками, на основе требований пользователя.
        /// </summary>
        /// <param name="activePage"> Страница, к которой применяется расчет </param>
        /// <param name="pageWidth"> Заполненная ячейками ширина страницы </param>
        /// <param name="pageHeight"> Заполненная ячейками высота страницы </param>
        /// <param name="cellWidth"> Ширина ячейки </param>
        /// <param name="cellHeight"> Высота ячейки </param>
        private void remap(Page activePage, out double pageWidth, out double pageHeight, double cellWidth, double cellHeight)
        {
            pageWidth = activePage.SizeWidth / cellWidth;
            pageWidth = Math.Truncate(pageWidth);
            pageWidth *= cellWidth;

            pageHeight = activePage.SizeHeight / cellHeight;
            pageHeight = Math.Truncate(pageHeight);
            pageHeight *= cellHeight;
        }
        #endregion
    }
}
