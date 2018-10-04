
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

public class R extends XWPFDocument
{

    /**
     * @param args
     */

    public  List<PInf> PS=new ArrayList<>();
    public  R (XWPFDocument doc) throws IOException, XmlException
    {



        CTStyle[] st = null;
        try
        {
            st = doc.getStyle().getStyleArray();
        } catch (XmlException | IOException e1)
        {
            // TODO 自动生成的 catch 块
            e1.printStackTrace();
        }


        CTSdtBlock[] block = doc.getDocument().getBody().getSdtArray();
        int a = 0;
        for (CTSdtBlock ctSdtBlock : block)
        {
            String font0 = ctSdtBlock.getSdtPr().getRPrArray(0).getRFonts().getAscii();
            int size0 = ctSdtBlock.getSdtPr().getRPrArray(0).getSz().getVal().intValue();

            XWPFSDT c = new XWPFSDT(ctSdtBlock, doc);
            List<XWPFParagraph> pas = c.getContent().getParagraphs();
            for (XWPFParagraph pa : pas)
            {
                //根据style   获取标题级数和字体
                String font1 = font0;
                int size1 = size0;
                String A=null;
                String tab=null;
                if (pa.getCTP().getPPr() != null)

                    if (pa.getCTP().getPPr().getPStyle() != null)

                        if (doc.getStyle().getStyleArray(find(st, pa.getCTP().getPPr().getPStyle().getVal())) != null)
                        {
                            CTStyle xxx = doc.getStyle().getStyleArray(find(st, pa.getCTP().getPPr().getPStyle().getVal()));
                            A = doc.getStyle().getStyleArray(find(st, pa.getCTP().getPPr().getPStyle().getVal())).getName().getVal();
                            //System.out.println(pa.getCTP().getPPr().getPStyle().getVal() + "  " + A + "  " + pa.getText());
                            if (doc.getStyle().getStyleArray(find(st, pa.getCTP().getPPr().getPStyle().getVal())).getRPr() != null)
                            {
                                
                                CTRPr rpr = doc.getStyle().getStyleArray(find(st, pa.getCTP().getPPr().getPStyle().getVal())).getRPr();
                                if (rpr.getSz() != null)
                                    size1 = rpr.getSz().getVal().intValue();
                                if (rpr.getRFonts() != null)
                                    if (rpr.getRFonts().getAscii() != null)
                                        font1 = rpr.getRFonts().getAscii();
                            }
                        }
                //最近风格
                if (pa.getCTP() != null)
                    if (pa.getCTP().getPPr() != null)
                        if (pa.getCTP().getPPr().getRPr() != null)
                        {
                            CTParaRPr rpr = pa.getCTP().getPPr().getRPr();
                            if (rpr.getRFonts() != null)
                                if (rpr.getRFonts().getAscii() != null)
                                    font1 = rpr.getRFonts().getAscii();

                            if (rpr.getSz() != null)
                                size1 = rpr.getSz().getVal().intValue();
                        }
                 //从目录里 获取tab的格式 leader
                if(pa.getCTP().getPPr()!=null)
                if (pa.getCTP().getPPr().getTabs()!=null){
                    if (pa.getCTP().getPPr().getTabs().getTabList()!=null){
                       tab=pa.getCTP().getPPr().getTabs().getTabList().get(0).getLeader().toString();

                    }

                }
                PInf P = new PInf(A, font1, size1,pa.getText(),tab);
                PS.add(P);

//                for (XWPFRun run : pa.getRuns())
//                {
//                    System.out.println(run);
//                    if (run.toString().length() != 0)
//                    {
//                        //根据style
//                        run.getCTR().getRPr().getRStyle();
//
//                        //最近风格
//                        run.getCTR().getRPr().getSz().getVal().intValue();
//                        run.getCTR().getRPr().getRFonts().getAscii();
//                    }
//                }

            }
        }
    }

     boolean isBT(XWPFParagraph p)
    {
        String v = null;
        char ch = ' ';
        if (p.getCTP() != null)
            if (p.getCTP().getPPr() != null)
                if (p.getCTP().getPPr().getPStyle() != null)
                {
                    v = p.getCTP().getPPr().getPStyle().getVal();
                    for (int i = 0; i < v.length(); i++)
                    {
                        ch = v.charAt(i);
                        if (ch >= 'A' && ch <= 'Z' || ch >= 'a' && ch <= 'z')
                            return false;
                    }
                    if (v != null)
                        return true;
                }

        return false;
    }

    public  int find(CTStyle[] st, String ID)
    {
        int i = 0;
        for (CTStyle ctStyle : st)
        {
            if (ctStyle.getStyleId().equals(ID))
            {
                return i;
            }
            i++;
        }
        return -1;

    }
}
