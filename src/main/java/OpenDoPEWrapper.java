import org.docx4j.XmlUtils;
import org.docx4j.customXmlProperties.DatastoreItem;
import org.docx4j.customXmlProperties.SchemaRefs;
import org.docx4j.model.datastorage.CustomXmlDataStorage;
import org.docx4j.model.datastorage.CustomXmlDataStorageImpl;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.CustomXmlDataStoragePropertiesPart;
import org.docx4j.openpackaging.parts.CustomXmlPart;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.opendope.ConditionsPart;
import org.docx4j.openpackaging.parts.opendope.XPathsPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.opendope.conditions.Condition;
import org.opendope.conditions.Conditions;
import org.opendope.conditions.Xpathref;
import org.opendope.xpaths.ObjectFactory;
import org.opendope.xpaths.Xpaths;

import java.util.Map;
import java.util.UUID;
import java.util.stream.Stream;


public class OpenDoPEWrapper {
    private WordprocessingMLPackage wordprocessingMLPackage;
    private ObjectFactory xpathsOF = new ObjectFactory();
    private org.opendope.conditions.ObjectFactory conditionsOF = new org.opendope.conditions.ObjectFactory();
    private Xpaths xPaths;
    private Conditions conditions;

    public OpenDoPEWrapper(WordprocessingMLPackage wordprocessingMLPackage) throws Docx4JException {
        this.wordprocessingMLPackage = wordprocessingMLPackage;
        if (getXpathPart() == null) {
            wordprocessingMLPackage.getMainDocumentPart().addTargetPart(
                    createXpathsPart(),
                    RelationshipsPart.AddPartBehaviour.RENAME_IF_NAME_EXISTS);
            addProperties(getXpathPart(), "");
        }
        if (getCustomXmlDataStorageParts().size() == 0) {
            createCustomXmlStoragePart();
        }
        if (getConditionsPart() == null) {
            wordprocessingMLPackage.getMainDocumentPart().addTargetPart(
                    createConditionsPart(),
                    RelationshipsPart.AddPartBehaviour.RENAME_IF_NAME_EXISTS);
            addProperties(getConditionsPart(), "http://opendope.org/conditions");
        }
    }

    public Xpaths.Xpath addXpath(String id, String storeItemId, String xPath) {
        if (xPaths.getXpath().stream().anyMatch(p -> p.getId().equals(id))) {
            throw new IllegalArgumentException("Path with id already exist");
        }

        Xpaths.Xpath xpath = xpathsOF.createXpathsXpath();
        Xpaths.Xpath.DataBinding dataBinding = new Xpaths.Xpath.DataBinding();

        dataBinding.setStoreItemID(storeItemId);
        dataBinding.setXpath(xPath);
        xpath.setDataBinding(dataBinding);
        xpath.setId(id);

        xPaths.getXpath().add(xpath);

        //getXpathPart().setContents(xPaths);

        return xpath;
    }

    public Xpaths.Xpath addXpath(String xPath) {
        String storeId = getCustomXmlDataStorageParts().entrySet().stream().findFirst().get().getKey();
        String pathId = "x" + xPaths.getXpath().size();
        return addXpath(pathId, storeId, xPath);
    }

    public Condition addCondition(String id, String xpathId) {
        if (xPaths.getXpath().stream().noneMatch(p -> p.getId().equals(xpathId))) {
            throw new IllegalArgumentException("Xpath with id " + id + " not found");
        }
        if (conditions.getCondition().stream().anyMatch(c -> c.getId().equals(id))) {
            throw new IllegalArgumentException("Condition with id " + id + " already exist");
        }

        Condition condition = conditionsOF.createCondition();
        Xpathref xpathref = conditionsOF.createXpathref();

        xpathref.setId(xpathId);
        condition.setId(id);
        condition.setParticle(xpathref);

        conditions.getCondition().add(condition);
        //getConditionsPart().setContents(conditions);

        return condition;
    }

    private void createCustomXmlStoragePart() throws Docx4JException {
        org.docx4j.openpackaging.parts.CustomXmlDataStoragePart customXmlDataStoragePart =
                new org.docx4j.openpackaging.parts.CustomXmlDataStoragePart();
        CustomXmlDataStorage data = new CustomXmlDataStorageImpl();
        org.w3c.dom.Document domDoc = XmlUtils.neww3cDomDocument();

        domDoc.appendChild(domDoc.createElement("repository"));
        data.setDocument(domDoc);
        customXmlDataStoragePart.setData(data);

        wordprocessingMLPackage.getMainDocumentPart().addTargetPart(customXmlDataStoragePart, RelationshipsPart.AddPartBehaviour.RENAME_IF_NAME_EXISTS);

        String id = addProperties(customXmlDataStoragePart);

        getCustomXmlDataStorageParts().put(id, customXmlDataStoragePart);
    }

    public static String addProperties(Part parentPart, String... ns) throws Docx4JException {
        CustomXmlDataStoragePropertiesPart part = new CustomXmlDataStoragePropertiesPart();
        org.docx4j.customXmlProperties.ObjectFactory of = new org.docx4j.customXmlProperties.ObjectFactory();
        DatastoreItem dsi = of.createDatastoreItem();
        String newItemId = "{" + UUID.randomUUID().toString() + "}";
        dsi.setItemID(newItemId);

        SchemaRefs srefs = of.createSchemaRefs();
        dsi.setSchemaRefs(srefs);
        Stream.of(ns).forEach(s -> {
            SchemaRefs.SchemaRef sref = of.createSchemaRefsSchemaRef();
            sref.setUri(s);
            srefs.getSchemaRef().add(sref);

        });

        part.setJaxbElement(dsi);

        parentPart.addTargetPart(part, RelationshipsPart.AddPartBehaviour.RENAME_IF_NAME_EXISTS);

        return newItemId;
    }

    private ConditionsPart createConditionsPart() throws Docx4JException {
        ConditionsPart conditionsPart = new ConditionsPart(new PartName("/customXml/item1.xml"));
        conditions = conditionsOF.createConditions();
        conditionsPart.setContents(conditions);

        return conditionsPart;
    }

    private XPathsPart createXpathsPart() throws Docx4JException {
        XPathsPart xPathsPart = new XPathsPart(new PartName("/customXml/item1.xml"));
        xPaths = xpathsOF.createXpaths();
        xPathsPart.setContents(xPaths);

        return xPathsPart;
    }

    public Map<String, CustomXmlPart> getCustomXmlDataStorageParts() {
        return wordprocessingMLPackage.getCustomXmlDataStorageParts();
    }

    public XPathsPart getXpathPart() {
        return wordprocessingMLPackage.getMainDocumentPart().getXPathsPart();
    }

    public ConditionsPart getConditionsPart() {
        return wordprocessingMLPackage.getMainDocumentPart().getConditionsPart();
    }

    public WordprocessingMLPackage getWordprocessingMLPackage() {
        return wordprocessingMLPackage;
    }
}
