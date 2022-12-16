# Practic_rep
Репозиторий, используемый во время практики.

Код исполльзуеммый во время практики:
using static System.Net.Mime.MediaTypeNames;
using System.Reflection;

object file = Path.GetDirectoryName(Application.ExecutablePath) + @"\Answer.doc";

Word.Application wordObject = new Word.ApplicationClass();
wordObject.Visible = false;

object nullobject = Missing.Value;
Word.Document docs = wordObject.Documents.Open
    (ref file, ref nullobject, ref nullobject, ref nullobject,
    ref nullobject, ref nullobject, ref nullobject, ref nullobject,
    ref nullobject, ref nullobject, ref nullobject, ref nullobject,
    ref nullobject, ref nullobject, ref nullobject, ref nullobject);

String strLine;
bool bolEOF = false;

docs.Characters[1].Select();

int index = 0;
do
{
    object unit = Word.WdUnits.wdLine;
    object count = 1;
    wordObject.Selection.MoveEnd(ref unit, ref count);

    strLine = wordObject.Selection.Text;
    richTextBox1.Text += ++index + " - " + strLine + "\r\n"; //for our understanding

    object direction = Word.WdCollapseDirection.wdCollapseEnd;
    wordObject.Selection.Collapse(ref direction);

    if (wordObject.Selection.Bookmarks.Exists(@"\EndOfDoc"))
        bolEOF = true;
} while (!bolEOF);

docs.Close(ref nullobject, ref nullobject, ref nullobject);
wordObject.Quit(ref nullobject, ref nullobject, ref nullobject);
docs = null;
wordObject = null;