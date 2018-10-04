import java.math.BigInteger;

import org.apache.poi.util.Internal;
import org.apache.poi.util.LocaleUtil;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute.Space;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTParaRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtBlock;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtContentBlock;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtEndPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabStop;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabs;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabTlc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTheme;


public class TOC1 {
    CTSdtBlock block;

    public TOC1() {
        this(CTSdtBlock.Factory.newInstance());
    }

    public TOC1(CTSdtBlock block) {
        this.block = block;//sdt
        CTSdtPr sdtPr = block.addNewSdtPr();//sdtpr
        CTDecimalNumber id = sdtPr.addNewId();//
        id.setVal(new BigInteger("4844945"));
        sdtPr.addNewDocPartObj().addNewDocPartGallery().setVal("Table of contents");
        CTSdtEndPr sdtEndPr = block.addNewSdtEndPr();//��������δ֪
        CTRPr rPr = sdtEndPr.addNewRPr();
        CTFonts fonts = rPr.addNewRFonts();//����
        fonts.setAsciiTheme(STTheme.MINOR_H_ANSI);//�����ֵ���������
        fonts.setEastAsiaTheme(STTheme.MINOR_H_ANSI);
        fonts.setHAnsiTheme(STTheme.MINOR_H_ANSI);
        fonts.setCstheme(STTheme.MINOR_BIDI);
        rPr.addNewB().setVal(STOnOff.OFF);//��������
        rPr.addNewBCs().setVal(STOnOff.OFF);
        rPr.addNewColor().setVal("auto");//��ɫ  ʹ��ʮ������
        rPr.addNewSz().setVal(new BigInteger("99"));//�����С ��λΪ1/2��
        rPr.addNewSzCs().setVal(new BigInteger("24"));
        CTSdtContentBlock content = block.addNewSdtContent();

    }

    @Internal
    public CTSdtBlock getBlock() {
        return this.block;
    }
    /*
     * 
     * ���һ��Ŀ¼  
     * style  Ŀ¼�����������Ŀ¼��ǰ������
     * title  Ŀ¼������
     * page   Ŀ¼��ҳ��
     * bookmarfRef  ��ǩ����ʵ��Ŀ¼�ĵ����ת
     * 
     * */

    public void addRow(String style, String title, int page, String bookmarkRef) {
        CTSdtContentBlock contentBlock = this.block.getSdtContent();
        CTP p = contentBlock.addNewP();
        p.setRsidR("00EF7E24".getBytes(LocaleUtil.CHARSET_1252));
        p.setRsidRDefault("00EF7E24".getBytes(LocaleUtil.CHARSET_1252));
        CTPPr pPr = p.addNewPPr();
        CTParaRPr rPr=pPr.addNewRPr();
        pPr.addNewInd().setLeft(BigInteger.valueOf((Integer.valueOf(style)-1)*460));//���ݴ����style��ֵ ��������  ÿһ����460
        CTSpacing spacing =pPr.addNewSpacing();
        spacing.setLineRule(STLineSpacingRule.EXACT);//����й��򣬴˴����������м��
        spacing.setLine(BigInteger.valueOf(460));//�м��Ϊ460.���ֵ�λδ֪
        rPr.addNewRStyle().setVal("a7");
        //rPr.addNewSz().setVal(BigInteger.valueOf(10));;
        if (style.equals("1")) {
			style="11";
		}//style����Ϊһ  �Ǳ���ķ�� ���ּӴּӴ� �м����    ���Խ�1�滻��11  
        pPr.addNewPStyle().setVal(style);
        CTTabs tabs = pPr.addNewTabs();
        CTTabStop tab = tabs.addNewTab();
        tab.setVal(STTabJc.RIGHT);
        tab.setLeader(STTabTlc.DOT);
        tab.setPos(new BigInteger("8290"));
        pPr.addNewRPr().addNewNoProof();
        CTR run = p.addNewR();
        CTRPr rpr=run.addNewRPr();
        rpr.addNewNoProof();

        rpr.addNewSz().setVal(new BigInteger("24"));
        rpr.addNewRFonts().setAscii("����");
        run.addNewT().setStringValue(title);
        run = p.addNewR();
        rpr=run.addNewRPr();
        rpr.addNewNoProof();
        rpr.addNewSz().setVal(new BigInteger("24"));
        run.addNewTab();
        run = p.addNewR();
        run.addNewRPr().addNewNoProof();
        run.addNewFldChar().setFldCharType(STFldCharType.BEGIN);
        // pageref run
        run = p.addNewR();
        run.addNewRPr().addNewNoProof();
        run.addNewRPr().addNewSz().setVal(new BigInteger("24"));
        
        CTText text = run.addNewInstrText();
        text.setSpace(Space.PRESERVE);
        // bookmark reference
        text.setStringValue(" PAGEREF _Toc" + bookmarkRef + " \\h ");
        p.addNewR().addNewRPr().addNewNoProof();
        run = p.addNewR();
        run.addNewRPr().addNewNoProof();
        run.addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
        // page number run
        run = p.addNewR();
        run.addNewRPr().addNewNoProof();
        run.addNewRPr().addNewSz().setVal(new BigInteger("24"));
        rpr.addNewRFonts().setAscii("����");
        run.addNewT().setStringValue(Integer.toString(page));
        run = p.addNewR();
        run.addNewRPr().addNewNoProof();//ƴд���
        run.addNewFldChar().setFldCharType(STFldCharType.END);
    }
}
