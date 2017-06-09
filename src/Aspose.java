
import java.util.regex.Pattern;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FindReplaceOptions;


public class Aspose 
{
    public static void main( String[] args ) throws Exception
    {
		try {
	        String MyDir="/mydata/";
//	        String MyDir="E:/workspace/Aspose/";
	        Document doc = new Document(MyDir+"input.docx");

	        FindReplaceOptions options = new FindReplaceOptions();
	        options.ReplacingCallback = new FindandInsertOLE(MyDir+"1.xlsx");
	        Pattern regex = Pattern.compile("\\{excel\\}", Pattern.CASE_INSENSITIVE);
	        doc.getRange().replace(regex, "", options);

	        // Save the output document.
	        doc.save(MyDir+"output.docx");
		} catch (Exception e) {
			e.printStackTrace();
		}
    }
}