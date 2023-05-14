import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.InputStream;

public class powerpointPresentation {


    public static void main(String[] args) throws Exception {
        // Creating a new .pptx file.
        XMLSlideShow ppt = new XMLSlideShow();
        // Setting the parameters for the first slide of the PowerPoint file.
        XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
        // Clarifying the slide layout. In this case it is the title slide.
        XSLFSlideLayout layout = defaultMaster.getLayout(SlideLayout.TITLE_ONLY);

        // Finally, creating the first slide, passing the layout from before.
        XSLFSlide slide = ppt.createSlide(layout);
        slide.getBackground().setFillColor(Color.DARK_GRAY);
        XSLFTextShape title = slide.getPlaceholder(0);
        // Removing the predefined text.
        title.clearText();

        // Creating a new paragraph.
        XSLFTextParagraph p = title.addNewTextParagraph();
        XSLFTextRun r = p.addNewTextRun();
        r.setText("Sage Presentation");
        r.setFontColor(Color.GREEN);
        r.setFontSize(50.);

        // Adding an image to the slide from maven's resources folder.
        InputStream is = powerpointPresentation.class.getResourceAsStream("/Presentation-Free-PNG-Image.png");
        byte[] pd = IOUtils.toByteArray(is);
        XSLFPictureData pictureData = ppt.addPicture(pd, PictureData.PictureType.PNG);
        // Defining image position.
        XSLFPictureShape pictureShape = slide.createPicture(pictureData);
        pictureShape.setAnchor(new Rectangle(20, 80, 680, 388));

        // Creating a second slide.
        layout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
        slide = ppt.createSlide(layout);
        title = slide.getPlaceholder(0);
        title.clearText();
        r = title.addNewTextParagraph().addNewTextRun();
        r.setText("Why I applied for this role");
        r.setFontColor(Color.GREEN);
        r.setFontSize(50.);

        // Adding paragraph.
        XSLFTextShape content = slide.getPlaceholder(1);
        content.clearText();
        XSLFTextParagraph p2 = content.addNewTextParagraph();
        XSLFTextRun r2 = p2.addNewTextRun();
        XSLFTextParagraph p3 = content.addNewTextParagraph();
        XSLFTextRun r3 = p3.addNewTextRun();
        XSLFTextParagraph p4 = content.addNewTextParagraph();
        XSLFTextRun r4 = p4.addNewTextRun();
        r2.setText("I am a recent graduate, looking to break into the software industry.");
        r3.setText("Passionate about Java and enjoy using it. Used it as part of an OOP module at university.");
        r4.setText("I am eager to work for Sage as they are a market leaders with a diverse clientele and tech stack. Background in Accounting.");
        r2.setFontSize(25.);
        r3.setFontSize(25.);
        r4.setFontSize(25.);
        r2.setFontColor(Color.GREEN);
        r3.setFontColor(Color.GREEN);
        r4.setFontColor(Color.GREEN);

        // Creating a third slide.
        layout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
        slide = ppt.createSlide(layout);
        title = slide.getPlaceholder(0);
        title.clearText();
        r = title.addNewTextParagraph().addNewTextRun();
        r.setText("What most interests me about this role");
        r.setFontColor(Color.GREEN);
        r.setFontSize(50.);

        // Adding paragraph.
        XSLFTextShape content2 = slide.getPlaceholder(1);
        content2.clearText();
        XSLFTextParagraph s = content2.addNewTextParagraph();
        XSLFTextRun t = s.addNewTextRun();
        XSLFTextParagraph s2 = content2.addNewTextParagraph();
        XSLFTextRun t2 = s2.addNewTextRun();
        XSLFTextParagraph s3 = content2.addNewTextParagraph();
        XSLFTextRun t3 = s3.addNewTextRun();
        t.setText("The mentoring and training provided.");
        t2.setText("The chance to work with a diverse tech stack, particularly AWS, and tackle diverse problems that are commonplace in larger companies.");
        t3.setText("Working as part of a multidisciplinary team to meet goals.");
        t.setFontSize(25.);
        t2.setFontSize(25.);
        t3.setFontSize(25.);
        t.setFontColor(Color.GREEN);
        t2.setFontColor(Color.GREEN);
        t3.setFontColor(Color.GREEN);

        // And so on.
        layout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
        slide = ppt.createSlide(layout);
        title = slide.getPlaceholder(0);
        title.clearText();
        r = title.addNewTextParagraph().addNewTextRun();
        r.setText("My long-term career ambitions");
        r.setFontColor(Color.GREEN);
        r.setFontSize(50.);

        // Adding paragraph.
        XSLFTextShape content3 = slide.getPlaceholder(1);
        content3.clearText();
        XSLFTextParagraph z = content3.addNewTextParagraph();
        XSLFTextRun f = z.addNewTextRun();
        XSLFTextParagraph z2 = content3.addNewTextParagraph();
        XSLFTextRun f2 = z2.addNewTextRun();
        XSLFTextParagraph z3 = content3.addNewTextParagraph();
        XSLFTextRun f3 = z3.addNewTextRun();
        f.setText("To delve deeper into machine learning and AI.");
        f2.setText("To that end, completing an MSc in AI in the near future.");
        f3.setText("Develop my leadership and management skills.");
        f.setFontSize(25.);
        f2.setFontSize(25.);
        f3.setFontSize(25.);
        f.setFontColor(Color.GREEN);
        f2.setFontColor(Color.GREEN);
        f3.setFontColor(Color.GREEN);

        layout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
        slide = ppt.createSlide(layout);
        title = slide.getPlaceholder(0);
        title.clearText();
        r = title.addNewTextParagraph().addNewTextRun();
        r.setText("What I want from Sage");
        r.setFontColor(Color.GREEN);
        r.setFontSize(50.);

        // Adding paragraph.
        XSLFTextShape content4 = slide.getPlaceholder(1);
        content4.clearText();
        XSLFTextParagraph v = content4.addNewTextParagraph();
        XSLFTextRun y = v.addNewTextRun();
        XSLFTextParagraph v2 = content4.addNewTextParagraph();
        XSLFTextRun y2 = v2.addNewTextRun();
        XSLFTextParagraph v3 = content4.addNewTextParagraph();
        XSLFTextRun y3 = v3.addNewTextRun();
        y.setText("The chance to grow and develop. With clear opportunities for career progression.");
        y2.setText("A chance to work on a variety of projects and with a variety of different teams within an ever-evolving developmental cycle.");
        y3.setText("Opportunity for hybrid and remote work and well meaning and thoughtful management.");
        y.setFontSize(25.);
        y2.setFontSize(25.);
        y3.setFontSize(25.);
        y.setFontColor(Color.GREEN);
        y2.setFontColor(Color.GREEN);
        y3.setFontColor(Color.GREEN);

        layout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
        slide = ppt.createSlide(layout);
        title = slide.getPlaceholder(0);
        title.clearText();
        r = title.addNewTextParagraph().addNewTextRun();
        r.setText("How I might be different from other candidates");
        r.setFontColor(Color.GREEN);
        r.setFontSize(50.);

        // Adding paragraph.
        XSLFTextShape content5 = slide.getPlaceholder(1);
        content5.clearText();
        XSLFTextParagraph i = content5.addNewTextParagraph();
        XSLFTextRun o = i.addNewTextRun();
        XSLFTextParagraph i2 = content5.addNewTextParagraph();
        XSLFTextRun o2 = i2.addNewTextRun();
        XSLFTextParagraph i3 = content5.addNewTextParagraph();
        XSLFTextRun o3 = i3.addNewTextRun();
        o.setText("I am perhaps a little older than the typical graduate and this have given me the chance to gain more professional experience.");
        o2.setText("Well acquainted with independent learning from my time at the Open University. Continued learning following graduation.");
        o3.setText("Also familiar with related technologies and just generally very curious and driven.");
        o.setFontSize(25.);
        o2.setFontSize(25.);
        o3.setFontSize(25.);
        o.setFontColor(Color.GREEN);
        o2.setFontColor(Color.GREEN);
        o3.setFontColor(Color.GREEN);

        // Saving presentation.
        FileOutputStream out = new FileOutputStream("SagePresentation.pptx");
        ppt.write(out);
        out.close();
        ppt.close();
    }

}


