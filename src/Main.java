import java.io.*;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtBlock;

public class Main {

    public static void main(String[] args) throws IOException {

        T t=null;
        R r=null;
        InputStream in=new FileInputStream("AYDS.docx");

        XWPFDocument docin=new XWPFDocument(in);


        try {
            t=new T(docin);
        } catch (XmlException e) {
            e.printStackTrace();
        }

        try {
            r=new R(docin);
        } catch (XmlException e) {
            e.printStackTrace();
        }
        List<MP> ts=t.getPs();
        List<PInf> Ps=r.PS;
        int i=0;
        for (MP s:ts) {
            String text1;
            int index=Ps.get(i).text.indexOf("\t");
            if (index!=-1)
            {text1=Ps.get(i).text.substring(0,index);}
            else {text1=Ps.get(i).text;}
           if (!s.getbT().equals(Ps.get(i).BT)||!s.getText().equals(text1)||!Ps.get(i).tab.equals("dot")){
               System.out.println(s.getbT()+"||\t||"+s.getText()+"||\t||"+Ps.get(i).BT+"||\t||"+text1+"|| ....类型是" + Ps.get(i).tab+"应是‘dot’");
           }


           i++;
        }
    }
}