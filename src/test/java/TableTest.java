import org.docx4j.Docx4J;
import org.docx4j.TraversalUtil;
import org.docx4j.finders.ClassFinder;
import org.docx4j.model.table.TblFactory;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.opendope.xpaths.Xpaths;

import java.io.InputStream;

public class TableTest {
    private InputStream xmlTestData;

    @Before
    public void prepare() {
        ClassLoader classLoader = getClass().getClassLoader();
        xmlTestData = classLoader.getResourceAsStream("generator-test-data.xml");
    }

    public static P createP(String value) {
        ObjectFactory of = new ObjectFactory();

        P p = of.createP();
        Text text = of.createText();
        text.setValue(value);

        R run = of.createR();
        run.getContent().add(text);

        p.getContent().add(run);

        return p;
    }

    @Test
    public void createRepeaterControlTest() throws Exception {
        OpenDoPEWrapper doPEWrapper = new OpenDoPEWrapper(WordprocessingMLPackage.createPackage());

        ContentControlFactory ccf = new ContentControlFactory();
        Tbl tbl = TblFactory.createTable(1, 1, 100);

        //headers

        Tr tr = new Tr();
        Tc tc = new Tc();

        tc.getContent().add(createP("apples"));

        tr.getContent().add(ccf.createContentControl(doPEWrapper.addXpath("/invoice[1]/items/item[1]/name"), tc));
        tbl.getContent().add(ccf.createRepeaterControl(doPEWrapper.addXpath("/invoice[1]/items/item"), tr));

        doPEWrapper.getWordprocessingMLPackage().getMainDocumentPart().getContent().add(tbl);

        Docx4J.bind(doPEWrapper.getWordprocessingMLPackage(), xmlTestData, Docx4J.FLAG_BIND_INSERT_XML | Docx4J.FLAG_BIND_BIND_XML);

        //Docx4J.save(doPEWrapper.getWordprocessingMLPackage(), new File("test-template.docx"));
//

        ClassFinder finder = new ClassFinder(Tr.class);

        new TraversalUtil(doPEWrapper.getWordprocessingMLPackage().getMainDocumentPart().getContent(), finder);

        Assert.assertEquals(4, finder.results.size());
    }
}
