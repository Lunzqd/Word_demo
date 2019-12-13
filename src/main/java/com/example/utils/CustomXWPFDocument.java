package com.example.utils;


import org.apache.poi.POIXMLException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlToken;
import org.apache.xmlbeans.impl.values.XmlAnyTypeImpl;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 自定义 XWPFDocument，并重写 createPicture()方法
 */
public class CustomXWPFDocument extends XWPFDocument {
    public CustomXWPFDocument(InputStream in) throws IOException {
        super(in);
    }

    public CustomXWPFDocument() {
        super();
    }

    public CustomXWPFDocument(OPCPackage pkg) throws IOException {
        super(pkg);
    }

    public void renderPicture(int id, int width, int height, XWPFParagraph paragraph, XWPFRun run) throws Exception {
        final int EMU = 9525;
        width *= EMU;
        height *= EMU;
        String blipId = getAllPictures().get(id).getPackageRelationship().getId();
        CTDrawing drawing = paragraph.createRun().getCTR().addNewDrawing();

        // 上下型环绕
        //wp:wrapTopAndBottom

        // 图片衬于文字下方
        String xml = "<wp:anchor allowOverlap=\"1\" layoutInCell=\"1\" locked=\"1\" behindDoc=\"1\" relativeHeight=\"0\" simplePos=\"0\" distR=\"114300\" distL=\"114300\" distB=\"0\" distT=\"0\" " +
                " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"" +
                " xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\"" +
                " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"  >" +
                "<wp:simplePos y=\"0\" x=\"0\"/>" +
                "<wp:positionH relativeFrom=\"column\">" +
                "<wp:align>left</wp:align>" +
                "</wp:positionH>" +
                "<wp:positionV relativeFrom=\"paragraph\">" +
                "<wp:posOffset>30</wp:posOffset>" +
                "</wp:positionV>" +
                "<wp:extent cy=\"" + height + "\" cx=\"" + width + "\"/>" +
                "<wp:effectExtent b=\"0\" r=\"0\" t=\"0\" l=\"0\"/>" +
                "<wp:wrapNone/>" +
                "<wp:docPr descr=\"Picture Alt\" name=\"Picture Hit\" id=\"0\"/>" +
                "<wp:cNvGraphicFramePr>" +
                "<a:graphicFrameLocks noChangeAspect=\"true\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" />" +
                "</wp:cNvGraphicFramePr>" +
                "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                "<pic:nvPicPr>" +
                "<pic:cNvPr name=\"Picture Hit\" id=\"1\"/>" +
                "<pic:cNvPicPr/>" +
                "</pic:nvPicPr>" +
                "<pic:blipFill>" +
                "<a:blip r:embed=\"" + blipId + "\"/>" +
                "<a:stretch>" +
                "<a:fillRect/>" +
                "</a:stretch>" +
                "</pic:blipFill>" +
                "<pic:spPr>" +
                "<a:xfrm>" +
                "<a:off y=\"0\" x=\"0\"/>" +
                "<a:ext cy=\"" + height + "\" cx=\"" + width + "\"/>" +
                "</a:xfrm>" +
                "<a:prstGeom prst=\"rect\">" +
                "<a:avLst/>" +
                "</a:prstGeom>" +
                "</pic:spPr>" +
                "</pic:pic>" +
                "</a:graphicData>" +
                "</a:graphic>" +
                "<wp14:sizeRelH relativeFrom=\"margin\">" +
                "<wp14:pctWidth>0</wp14:pctWidth>" +
                "</wp14:sizeRelH>" +
                "<wp14:sizeRelV relativeFrom=\"margin\">" +
                "<wp14:pctHeight>0</wp14:pctHeight>" +
                "</wp14:sizeRelV>" +
                "</wp:anchor>";

        drawing.set(XmlToken.Factory.parse(xml, DEFAULT_XML_OPTIONS));
        CTPicture pic = getCTPictures(drawing).get(0);
        XWPFPicture xwpfPicture = new XWPFPicture(pic, paragraph);
        run.getEmbeddedPictures().add(xwpfPicture);
    }


    public static List<CTPicture> getCTPictures(XmlObject o) {
        List<CTPicture> pictures = new ArrayList<>();
        XmlObject[] picts = o.selectPath("declare namespace pic='"
                + CTPicture.type.getName().getNamespaceURI() + "' .//pic:pic");
        for (XmlObject pict : picts) {
            if (pict instanceof XmlAnyTypeImpl) {
                // Pesky XmlBeans bug - see Bugzilla #49934
                try {
                    pict = CTPicture.Factory.parse(pict.toString(),
                            DEFAULT_XML_OPTIONS);
                } catch (XmlException e) {
                    throw new POIXMLException(e);
                }
            }
            if (pict instanceof CTPicture) {
                pictures.add((CTPicture) pict);
            }
        }
        return pictures;
    }


}