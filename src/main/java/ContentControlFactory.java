import org.docx4j.wml.*;
import org.opendope.conditions.Condition;
import org.opendope.xpaths.Xpaths;

import static org.docx4j.jaxb.Context.getWmlObjectFactory;

public class ContentControlFactory {
    ObjectFactory wmlObjectFactory = getWmlObjectFactory();

    public SdtPr createSdtPr(CTDataBinding dataBinding, String tag, String alias) {
        SdtPr sdtPr = createSdtPr(tag, alias);
        sdtPr.setDataBinding(dataBinding);

        return sdtPr;
    }

    public SdtPr createSdtPr(String tag, String alias) {
        Tag t = new Tag();

        SdtPr.Alias a = new SdtPr.Alias();
        a.setVal(alias);
        t.setVal(tag);
        SdtPr sdtPr = wmlObjectFactory.createSdtPr();
        sdtPr.getRPrOrAliasOrLock().add(a);
        sdtPr.setTag(t);

        sdtPr.setId();
        return sdtPr;
    }

    public SdtBlock createRepeaterControl(Xpaths.Xpath repeaterXpath, Xpaths.Xpath xpath, Object template) {

        return createSdt(
                createSdtPr("od:repeat=" + repeaterXpath.getId(), "Repeat"),
                createSdtContent(createContentControl(xpath, template))
        );
    }

    public SdtBlock createRepeaterControl(Xpaths.Xpath repeaterXpath, Object template) {
        return createSdt(
                createSdtPr("od:repeat=" + repeaterXpath.getId(), "Repeat"),
                createSdtContent(template)
        );
    }

    public SdtBlock createContentControl(Xpaths.Xpath xpath, Object template) {
        return createSdt(
                createSdtPr(
                        createDataBinding(xpath.getDataBinding().getXpath(), xpath.getDataBinding().getStoreItemID()),
                        "od:xpath=" + xpath.getId(), ""),
                createSdtContent(template));
    }

    public SdtContentBlock createSdtContent(Object template) {
        SdtContentBlock sdtContent = wmlObjectFactory.createSdtContentBlock();

        sdtContent.getContent().add(template);
        return sdtContent;
    }

    public SdtBlock createConditionControl(Condition condition, Object template) {
        return createSdt(
                createSdtPr("od:condition=" + condition.getId(), ""),
                createSdtContent(template)
        );
    }

    public SdtBlock createSdt(SdtPr sdtPr, SdtContentBlock sdtContentBlock) {
        SdtBlock sdtBlock = wmlObjectFactory.createSdtBlock();
        sdtBlock.setSdtPr(sdtPr);
        sdtBlock.setSdtContent(sdtContentBlock);

        return sdtBlock;
    }

    public CTDataBinding createDataBinding(String xPath, String storeItemID) {
        CTDataBinding dataBinding = wmlObjectFactory.createCTDataBinding();
        dataBinding.setXpath(xPath);//xPath - это строка с XPath до XML-элемента, связанного с этим Content Control
        dataBinding.setStoreItemID(storeItemID);//storeItemID - это ID корневого XML-элемента, из которого нужно брать данные
        return dataBinding;
    }

}

