using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Activities;
using System.ComponentModel;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;

namespace DocTranslate
{

    public sealed class ReplaceTexts : CodeActivity
    {

        [RequiredArgument]
        [Category("Input")]
        public InArgument<string> Filename { get; set; }


        [RequiredArgument]
        [Category("Input")]
        public InArgument<Dictionary<string, string>> Texts { get; set; }


        // 작업 결과 값을 반환할 경우 CodeActivity<TResult>에서 파생되고
        // Execute 메서드에서 값을 반환합니다.
        protected override void Execute(CodeActivityContext context)
        {
            string filename = context.GetValue(this.Filename);
            Dictionary<string, string> dicText = context.GetValue(this.Texts);

            var app = new Application();
            var presentations = app.Presentations;
            Presentation objPres = presentations.Open(filename, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoFalse);
            Slides slides = objPres.Slides;
            if (slides != null)
            {
                int slide_count = 1;
                do
                {
                    var slide = slides[slide_count];
                    foreach (Microsoft.Office.Interop.PowerPoint.Shape textShape in slide.Shapes)
                    {
                        if (textShape.HasTextFrame == MsoTriState.msoTrue &&
                                 textShape.TextFrame.HasText == MsoTriState.msoTrue)
                        {
                            Microsoft.Office.Interop.PowerPoint.TextRange pptTextRange = textShape.TextFrame.TextRange;
                            if (pptTextRange != null && pptTextRange.Length > 0)
                            {
                                var id = string.Format("{0}.{1}", slide_count, textShape.Id);
                                if( dicText.ContainsKey( id))
                                {
                                    pptTextRange.Text = dicText[id].ToString();
                                }
                                Marshal.ReleaseComObject(pptTextRange);
                            }
                        }
                        Marshal.ReleaseComObject(textShape);
                    }
                    slide_count++;
                } while (slide_count <= slides.Count);
            }
            objPres.Save();
            app.Quit();

        }
    }
}
