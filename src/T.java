import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHpsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabStop;

public class T
{
    public List<MP> getPs() {
        return ps;
    }

    public void setPs(List<MP> ps) {
        this.ps = ps;
    }

    private List<MP> ps;
	 T(XWPFDocument doc) throws XmlException, IOException
	{

		List<XWPFParagraph> paras = doc.getParagraphs();
		CTStyle[] st = null;
		try
		{
			st = doc.getStyle().getStyleArray();
		} catch (XmlException | IOException e1)
		{
			// TODO 自动生成的 catch 块
			e1.printStackTrace();
		}
		BigInteger max = new BigInteger("0");
        for (XWPFParagraph para : paras)
        {
            if (para.getCTP().getPPr() != null)
                if (para.getCTP().getPPr().getPStyle() != null)
                    if (doc.getStyle().getStyleArray(find(st, para.getCTP().getPPr().getPStyle().getVal())) != null)
                        if (doc.getStyle().getStyleArray(find(st, para.getCTP().getPPr().getPStyle().getVal()))
                                .getRPr() != null)
                        {
                            if (doc.getStyle().getStyleArray(find(st, para.getCTP().getPPr().getPStyle().getVal()))
                                    .getRPr().getSz() != null) // getSzCs()
                            {
                                if (doc.getStyle().getStyleArray(find(st, para.getCTP().getPPr().getPStyle().getVal()))
                                        .getRPr().getSz().getVal().compareTo(max) == 1)
                                    max = doc.getStyle()
                                            .getStyleArray(find(st, para.getCTP().getPPr().getPStyle().getVal()))
                                            .getRPr().getSz().getVal();
                                // System.out.println(para.getText()+"\n\t"+doc.getStyle()
                                // .getStyleArray(find(st, para.getCTP().getPPr().getPStyle().getVal()))
                                // .getRPr().getSzCs().getVal());
                            }
                        }
        }


		 ps = new ArrayList<>();


		for (XWPFParagraph para : paras)
		{
			MP mp = new MP();
			if (para.getCTP().getPPr() != null)
			{
				boolean Is = findInSAndNew(doc, st, max, para, mp, ps);
				// 以下处理不是标题的情况 包含在各个rpr和ppr里的对应伪标题 手动填加
				if (!Is)
				{
					findInRAndNew(para, max, mp, ps);

				}
			} // else里面字符为空


		}
		for (MP p : ps) // 生成的List检测
		{
			if (p.getText().substring(0,2).equals("摘要"))
			    p.setText("摘要");
		}
	}
//从正文检测里面是否有学生学写的伪标题
	private static void findInRAndNew(XWPFParagraph para, BigInteger max, MP mp, List<MP> ps) {
        List<XWPFRun> rs = para.getRuns();
        for (XWPFRun r : rs) {
            // System.out.println(r);
            // System.out.println((int) r.getFontSize());
            //<editor-fold desc="一级标题if">
            if (r.getFontSize() * 2 == max.intValue())        //一级
            {
                if (r.getCTR() != null)
                    if (r.getCTR().getRPr() != null)
                        if (r.getCTR().getRPr().getRFonts() != null)
                            if (r.getCTR().getRPr().getRFonts().getAscii().equals("Times New Roman"))
                                continue;

                mp = new MP();
                mp.setbT(1);
                mp.setText(r.toString());
                mp.addPs(para);
                mp.setSzCs(new BigInteger(max.toString()));
                mp.setEastAsia("黑体");
                if (r.toString().length() != 0)
                    if (!r.toString().equals(" ") && !r.toString().equals("  ")
                            && !r.toString().equals("   "))
                        ps.add(mp);
            }
            //</editor-fold>
            //<editor-fold desc="二级标题if">
            if (r.getCTR() != null)                            //二级
                if (r.getCTR().getRPr() != null)
                    if (r.getCTR().getRPr().getRFonts() != null)
                        if (r.getCTR().getRPr().getRFonts().getEastAsia() != null)
                            if (r.getCTR().getRPr().getRFonts().getEastAsia().equals("黑体")
                                    && (r.getFontSize() * 2 != max.intValue()))
                                if ((r.toString().length() >= 3))
                                    if (!((r.toString().charAt(0) == '图') && (('0' <= (r.toString().charAt(1))
                                            && (r.toString().charAt(1)) <= '9')
                                            || ('0' <= (r.toString().charAt(2)) && (r.toString().charAt(2)) <= '9')))) // System.out.println((para.getText().charAt(0)=='图')&&(('0'<=(para.getText().charAt(1))&&(para.getText().charAt(1))<='9')||('0'<=(para.getText().charAt(2))&&(para.getText().charAt(2))<='9')));

                                        if ((r.toString().charAt(0) >= '0' && r.toString().charAt(0) <= '9')
                                                && (r.toString().charAt(2) >= '0' && r.toString().charAt(2) <= '9')
                                                && r.toString().charAt(1) == '.') {
                                            mp = new MP();
                                            mp.setbT(2);
                                            mp.setText(r.toString());
                                            mp.addPs(para);
                                            mp.setEastAsia("黑体");
                                            if (!r.toString().equals(" ") && !r.toString().equals("  ")
                                                    && !r.toString().equals("   "))
                                                ps.add(mp);

                                        }
            //</editor-fold>
            //<editor-fold desc="三级标题if">
            if (r.getCTR() != null)                            //三级
                if (r.getCTR().getRPr() != null)
                    if (r.getCTR().getRPr().getRFonts() != null)
                        if (r.getCTR().getRPr().getRFonts().getEastAsia() != null)
                            if (r.getCTR().getRPr().getRFonts().getEastAsia().equals("仿宋")
                                    && (r.getFontSize() * 2 != max.intValue())) {
                                mp = new MP();
                                mp.setbT(3);
                                mp.setText(r.toString());
                                mp.addPs(para);
                                if (r.toString().length() != 0)
                                    if (!r.toString().equals(" ") && !r.toString().equals("  ")
                                            && !r.toString().equals("   "))
                                        ps.add(mp);
                            }
            //</editor-fold>
            if (r.getCTR() != null)
                if (r.getCTR().getRPr() != null)
                    if (r.getCTR().getRPr().getSz() != null)
                        if (r.getCTR().getRPr().getSz().getVal().intValue() == 36) {


                            if (r.getCTR().getRPr().getRFonts() != null)
                                if (r.getCTR().getRPr().getRFonts().getAscii()!=null)
                                if (r.getCTR().getRPr().getRFonts().getAscii().equals("Times New Roman"))     
                                    continue;                                                                 
                            mp = new MP();
                            mp.setbT(1);
                            mp.setText(r.toString());
                            mp.addPs(para);
                            mp.setSzCs(new BigInteger("36"));
                            mp.setEastAsia("黑体");
                            if (r.toString().length() != 0)
                                if (!r.toString().equals(" ") && !r.toString().equals("  ")
                                        && !r.toString().equals("   "))
                                    ps.add(mp);
                        }
        }

    }

	// 从Style下面提取信息 从Style中找到并且建和添加到list 在if (para.getCTP().getPPr()
	// !=null)下调用建立在最大是一级 黑体是二级 仿宋是三级
	private static boolean findInSAndNew(XWPFDocument doc, CTStyle[] st, BigInteger max, XWPFParagraph para, MP mp,
										 List<MP> ps) throws XmlException, IOException

	{   if (para.getRuns().size()!=0)if (para.getRuns().get(0)!=null)if (para.getRuns().get(0).getCTR()!=null)if (para.getRuns().get(0).getCTR().getTArray().length!=0)
		if(para.getRuns().get(0).getCTR().getTArray(0).toString().equals("摘要")){

		mp = new MP();
		mp.setbT(1);
		mp.setText("摘要");
		mp.addPs(para);
		mp.setSzCs(new BigInteger(max.toString()));
		mp.setEastAsia("黑体");
		if (para.getText().length() != 0)
			if (!para.getText().equals(" ") && !para.getText().equals("  ")
					&& !para.getText().equals("   "))
				ps.add(mp);
		return true;
	}
		if (para.getCTP()!=null)if (para.getCTP().getPPr()!=null)if (para.getCTP().getPPr().getRPr()!=null)if (para.getCTP().getPPr().getRPr().getRFonts()!=null)if (para.getCTP().getPPr().getRPr().getRFonts().getAscii()!=null)if (para.getCTP().getPPr().getRPr().getRFonts().getAscii().equals("Times New Roman")) return false;

		if (para.getCTP().getPPr().getPStyle() != null)
		{

			if (doc.getStyle().getStyleArray(find(st, para.getCTP().getPPr().getPStyle().getVal())) != null)
				if (doc.getStyle().getStyleArray(find(st, para.getCTP().getPPr().getPStyle().getVal()))
                        .getRPr() != null)
                {
					CTHpsMeasure ls = doc.getStyle()
							.getStyleArray(find(st, para.getCTP().getPPr().getPStyle().getVal())).getRPr().getSz(); // 获取字号
					CTFonts lf = doc.getStyle().getStyleArray(find(st, para.getCTP().getPPr().getPStyle().getVal()))
							.getRPr().getRFonts(); // 获取字体
					if (lf != null)
					{
						if (lf.getEastAsia() != null)
						{



							if (lf.getEastAsia().equals("黑体")
									&& !(ls == null || ls.getVal().compareTo(max) != 0)) // 一级
							{
								mp = new MP();
								mp.setbT(1);
								mp.setText(para.getText());
								mp.addPs(para);
								mp.setSzCs(new BigInteger(max.toString()));
								mp.setEastAsia("黑体");
								if (para.getText().length() != 0)
									if (!para.getText().equals(" ") && !para.getText().equals("  ")
											&& !para.getText().equals("   "))
										ps.add(mp);
								return true;
							}
							if (lf.getEastAsia().equals("黑体")
									&& (ls == null || ls.getVal().compareTo(max) != 0)) // 二级
							{
								mp = new MP();
								mp.setbT(2);
								mp.setText(para.getText());
								mp.addPs(para);
								mp.setEastAsia("黑体");
								if ((para.getText().length() != 0) && !((para.getText().charAt(0) == '图')
										&& (('0' <= (para.getText().charAt(1)) && (para.getText().charAt(1)) <= '9')
												|| ('0' <= (para.getText().charAt(2))
														&& (para.getText().charAt(2)) <= '9'))))
								{
									// System.out.println((para.getText().charAt(0)=='图')&&(('0'<=(para.getText().charAt(1))&&(para.getText().charAt(1))<='9')||('0'<=(para.getText().charAt(2))&&(para.getText().charAt(2))<='9')));
									if (!para.getText().equals(" ") && !para.getText().equals("  ")
											&& !para.getText().equals("   "))
										ps.add(mp);
								}
								return true;
							}
							if ((lf.getEastAsia().equals("仿宋"))
									&& (ls == null || ls.getVal().compareTo(max) != 0))




							{
								mp = new MP();
								mp.setbT(3);
								mp.setText(para.getText());
								mp.addPs(para);
								if (para.getText().length() != 0)
									if (!para.getText().equals(" ") && !para.getText().equals("  ")
											&& !para.getText().equals("   "))
										ps.add(mp);
								return true;
							}
						}
					}
				}
		}
		//endregion
		return false;
	}

	private static int find(CTStyle[] st, String ID) // 辅助函数 在style中查找
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
