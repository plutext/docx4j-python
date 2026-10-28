import java.io.File;
import org.docx4j.markdown.MarkdownExporter;
import org.docx4j.model.text.TextOptions;
import org.docx4j.model.text.Texts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/** docx4j's Markdown of a file's text model, fields as given. */
public class ProbeMd {
	public static void main(String[] args) throws Exception {
		WordprocessingMLPackage pkg = WordprocessingMLPackage.load(new File(args[0]));
		TextOptions options = new TextOptions().set("fields", args[1]);
		System.out.print(new MarkdownExporter().export(Texts.of(pkg, options), pkg));
	}
}
