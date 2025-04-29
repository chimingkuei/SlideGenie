using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SlideGenie
{
    public class ImagePlacement
    {
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }

        public ImagePlacement(float left, float top, float width, float height)
        {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }
    }

    class Core
    {
        Application pptApp = new Application();

        private void SlideTitle(Slide slide, string titleText)
        {
            float titleLeft = 66;
            float titleTop = 58;
            float titleWidth = 828;
            float titleHeight = 104;
            var titleBox = slide.Shapes.AddTextbox(
                Orientation: MsoTextOrientation.msoTextOrientationHorizontal,
                Left: titleLeft,
                Top: titleTop,
                Width: titleWidth,
                Height: titleHeight
            );
            titleBox.TextFrame.AutoSize = (PpAutoSize)MsoTriState.msoFalse;
            titleBox.TextFrame.TextRange.Text = titleText;
            titleBox.TextFrame.TextRange.Font.Size = 48;
            titleBox.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
            titleBox.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
            titleBox.TextFrame.TextRange.Font.Name = "Microsoft JhengHei";
        }

        private void SlideImage(Slide slide, string imgPath, ImagePlacement placement)
        {
            slide.Shapes.AddPicture(
                FileName: imgPath,
                LinkToFile: MsoTriState.msoFalse,
                SaveWithDocument: MsoTriState.msoTrue,
                Left: placement.Left,
                Top: placement.Top,
                Width: placement.Width,
                Height: placement.Height
            );
        }

        /// <summary>
        /// 1cm = 約28.35pt
        /// </summary>
        /// <param name="titles"></param>
        /// <param name="imagePaths"></param>
        /// <param name="savePath"></param>
        public void BuildSlide(string[] titles, string[][] imageGroups, string savePath)
        {
            Presentation presentation = pptApp.Presentations.Add(MsoTriState.msoTrue);
            for (int i = 0; i < titles.Length; i++)
            {
                string titleText = titles[i];
                string[] images = imageGroups[i];
                // 新增空白頁
                var slide = presentation.Slides.Add(presentation.Slides.Count + 1, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank);
                SlideTitle(slide, titleText);
                // 根據圖片數量配置位置（此處以兩張圖片為例）
                var placements = new ImagePlacement[]
                {
                    new ImagePlacement(100f, 150f, 360f, 270f),  // 左圖
                    new ImagePlacement(520f, 150f, 360f, 270f)   // 右圖
                };
                for (int j = 0; j < images.Length && j < placements.Length; j++)
                {
                    SlideImage(slide, images[j], placements[j]);
                }
            }
            presentation.SaveAs(savePath);
            presentation.Close();
            pptApp.Quit();
        }

    }
}
