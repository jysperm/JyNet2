<%@ Page Language="C#"%>
<script runat="server">
    protected void Page_Load(object sender, EventArgs e)
    {
        Random Rnd = new Random();
        string strNum = Convert.ToString((char)Rnd.Next(67, 90)) + Convert.ToString((char)Rnd.Next(67, 90)) + Convert.ToString((char)Rnd.Next(67, 90)) + Convert.ToString((char)Rnd.Next(67, 90));
        string[] strFontName = new string[] {"宋体", "华文行楷", "华文隶书", "华文彩云", "方正舒体", "华文琥珀" };

        Session["VCode"] = strNum;
        
        System.Drawing.Color bgColor =System.Drawing.Color.FromArgb(Rnd.Next(0, 255),Rnd.Next(0, 255),Rnd.Next(0, 255));
        System.Drawing.Color foreColor = System.Drawing.Color.FromArgb(Rnd.Next(0, 255), Rnd.Next(0, 255), Rnd.Next(0, 255));
        System.Drawing.Font foreFont = new System.Drawing.Font(strFontName[Rnd.Next(0, 5)], Rnd.Next(11, 15), System.Drawing.FontStyle.Bold);
        System.Drawing.Bitmap Pic = new System.Drawing.Bitmap(80, 30, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(Pic);
        System.Drawing.Rectangle r = new System.Drawing.Rectangle(0, 0, 80, 30);
        g.FillRectangle(new System.Drawing.SolidBrush(bgColor), r);
        
        g.DrawString(strNum, foreFont, new System.Drawing.SolidBrush(foreColor), 2, 2);
        System.IO.MemoryStream mStream = new System.IO.MemoryStream();
        Pic.Save(mStream, System.Drawing.Imaging.ImageFormat.Gif);
        g.Dispose();
        Pic.Dispose();
        Response.ClearContent();
        Response.ContentType = "image/GIF";
        Response.BinaryWrite(mStream.ToArray());
        Response.End();
    }
</script>
