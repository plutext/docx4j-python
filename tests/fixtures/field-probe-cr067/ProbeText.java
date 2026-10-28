import java.io.File;
import java.time.ZoneId;
import org.docx4j.model.text.TextOptions;
import org.docx4j.model.text.Texts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/** docx4j's text model over a file, fields: update or cached; args: path, fields, zone or "-". */
public class ProbeText {
	public static void main(String[] args) throws Exception {
		WordprocessingMLPackage pkg = WordprocessingMLPackage.load(new File(args[0]));
		if (args.length > 2 && !"-".equals(args[2])) pkg.setDateTimeZone(ZoneId.of(args[2]));
		TextOptions options = new TextOptions().set("fields", args[1]);
		System.out.print(Texts.plainText(pkg, options));
	}
}
