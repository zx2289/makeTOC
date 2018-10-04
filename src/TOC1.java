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
        CTSdtEndPr sdtEndPr = block.addNewSdtEndPr();//具体作用未知
        CTRPr rPr = sdtEndPr.addNewRPr();
        CTFonts fonts = rPr.addNewRFonts();//字体
        fonts.setAsciiTheme(STTheme.MINOR_H_ANSI);//各种字的字体设置
        fonts.setEastAsiaTheme(STTheme.MINOR_H_ANSI);
        fonts.setHAnsiTheme(STTheme.MINOR_H_ANSI);
        fonts.setCstheme(STTheme.MINOR_BIDI);
        rPr.addNewB().setVal(STOnOff.OFF);//粗体设置
        rPr.addNewBCs().setVal(STOnOff.OFF);
        rPr.addNewColor().setVal("auto");//颜色  使用十六进制
        rPr.addNewSz().setVal(new BigInteger("99"));//字体大小 单位为1/2磅
        rPr.addNewSzCs().setVal(new BigInteger("24"));
        CTSdtContentBlock content = block.addNewSdtContent();

    }

    @Internal
    public CTSdtBlock getBlock() {
        return this.block;
    }
    /*
     * 
     * 添加一行目录  
     * style  目录风格用来控制目录的前的缩进
     * title  目录的文字
     * page   目录的页码
     * bookmarfRef  书签用来实现目录的点击跳转
     * 
     * */

    public void addRow(String style, String title, int page, String bookmarkRef) {
        CTSdtContentBlock contentBlock = this.block.getSdtContent();
        CTP p = contentBlock.addNewP();
        p.setRsidR("00EF7E24".getBytes(LocaleUtil.CHARSET_1252));
        p.setRsidRDefault("00EF7E24".getBytes(LocaleUtil.CHARSET_1252));
        CTPPr pPr = p.addNewPPr();
        CTParaRPr rPr=pPr.addNewRPr();
        pPr.addNewInd().setLeft(BigInteger.valueOf((Integer.valueOf(style)-1)*460));//根据传入的style数值 控制缩进  每一级差460
        CTSpacing spacing =pPr.addNewSpacing();
        spacing.setLineRule(STLineSpacingRule.EXACT);//添加行规则，此处用来设置行间距
        spacing.setLine(BigInteger.valueOf(460));//行间距为460.数字单位未知
        rPr.addNewRStyle().setVal("a7");
        //rPr.addNewSz().setVal(BigInteger.valueOf(10));;
        if (style.equals("1")) {
			style="11";
		}//style设置为一  是标题的风格 文字加粗加大 行间距变大    所以将1替换成11  
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
        rpr.addNewRFonts().setAscii("宋体");
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
        rpr.addNewRFonts().setAscii("宋体");
        run.addNewT().setStringValue(Integer.toString(page));
        run = p.addNewR();
        run.addNewRPr().addNewNoProof();//拼写检查
        run.addNewFldChar().setFldCharType(STFldCharType.END);
    }
}
