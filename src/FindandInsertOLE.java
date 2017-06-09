
import java.util.ArrayList;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.Run;


class FindandInsertOLE implements IReplacingCallback {
    String olepath;
    public FindandInsertOLE(String path)
    {
        olepath = path;
    }

    public int replacing(ReplacingArgs e) throws Exception {
        // This is a Run node that contains either the beginning or the complete match.
        Node currentNode = e.getMatchNode();

        // The first (and may be the only) run can contain text before the match,
        // in this case it is necessary to split the run.
        if (e.getMatchOffset() > 0)
            currentNode = splitRun((Run) currentNode, e.getMatchOffset());

        ArrayList runs = new ArrayList();

        // Find all runs that contain parts of the match string.
        int remainingLength = e.getMatch().group().length();
        while ((remainingLength > 0) && (currentNode != null) && (currentNode.getText().length() <= remainingLength)) {
            runs.add(currentNode);
            remainingLength = remainingLength - currentNode.getText().length();

            // Select the next Run node.
            // Have to loop because there could be other nodes such as BookmarkStart etc.
            do {
                currentNode = currentNode.getNextSibling();
            } while ((currentNode != null) && (currentNode.getNodeType() != NodeType.RUN));
        }

        // Split the last run that contains the match if there is any text left.
        if ((currentNode != null) && (remainingLength > 0)) {
            splitRun((Run) currentNode, remainingLength);
            runs.add(currentNode);
        }
        

        DocumentBuilder builder = new DocumentBuilder((Document)e.getMatchNode().getDocument());
        builder.moveTo((Run)runs.get(0));
        builder.insertOleObject(olepath, "Excel.Sheet.12", false, true, null);
        // Remove matched text
        for (Run run : (Iterable<Run>) runs)
            run.remove();

        // Signal to the replace engine to do nothing because we have already done all what we wanted.
        return ReplaceAction.SKIP;
    }

    /**
     * Splits text of the specified run into two runs. Inserts the new run just
     * after the specified run.
     */
    private Run splitRun(Run run, int position) throws Exception {
        Run afterRun = (Run) run.deepClone(true);
        afterRun.setText(run.getText().substring(position));
        run.setText(run.getText().substring((0), (0) + (position)));
        run.getParentNode().insertAfter(afterRun, run);
        return afterRun;
    }
}