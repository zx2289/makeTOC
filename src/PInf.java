import java.math.BigInteger;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;

public class PInf
{
	public String BT;
	public String font;
	public int size;
	public String text;
	public String tab;
	public PInf(String BT,String font,int size,String text,String tab)
	{
		this.size=size;
		this.BT=BT;
		this.font=font;
		this.text=text;
		this.tab=tab;
	}
}
