# spire.presentation javascript hello world
## create a PowerPoint presentation with Hello World text
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Add a new shape to the PPT document
let rec = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 250,80,(500 + ppt.SlideSize.Size.Width / 2 - 250),230);
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle:rec});

shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
shape.Fill.FillType = wasmModule.FillFormatType.None;

// Add text to the shape
shape.AppendTextFrame("Hello World!");

// Set the font and fill style of the text
let textRange = shape.TextFrame.TextRange;
textRange.Fill.FillType = wasmModule.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
textRange.FontHeight = 66;
textRange.LatinFont = wasmModule.TextFont;
```

---

# spire.presentation javascript paragraph
## add paragraph to PowerPoint slide
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Create a rectangle from the left, top, right, and bottom coordinates based on slide size
let rec = wasmModule.RectangleF.FromLTRB(0,0,ppt.SlideSize.Size.Width,ppt.SlideSize.Size.Height);

// Append an embedded image as a rectangle shape to the first slide
ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName:imageName,rectangle:rec});

// Set the line color of the first shape to Floral White
ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

// Append a new shape
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle:wasmModule.RectangleF.FromLTRB(50, 70, 670, 220)});
shape.Fill.FillType = wasmModule.FillFormatType.None;
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

// Set the alignment of paragraph
shape.TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Left;

// Set the indent of paragraph
shape.TextFrame.Paragraphs.get_Item(0).Indent = 50;

// Set the linespacing of paragraph
shape.TextFrame.Paragraphs.get_Item(0).LineSpacing = 150;

// Set the text of paragraph
shape.TextFrame.Text = "This powerful component suite contains the most up-to-date versions of all .NET WPF Silverlight components offered by E-iceblue.";

// Set the Font
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Arial Rounded MT Bold");
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_Black();
```

---

# PowerPoint Text Alignment
## Demonstrates how to set different text alignment types for paragraphs in a PowerPoint slide
```javascript
// Get the related shape and set the text alignment
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(1);
shape.TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Left;
shape.TextFrame.Paragraphs.get_Item(1).Alignment = wasmModule.TextAlignmentType.Center;
shape.TextFrame.Paragraphs.get_Item(2).Alignment = wasmModule.TextAlignmentType.Right;
shape.TextFrame.Paragraphs.get_Item(3).Alignment = wasmModule.TextAlignmentType.Justify;
shape.TextFrame.Paragraphs.get_Item(4).Alignment = wasmModule.TextAlignmentType.None;
```

---

# spire.presentation javascript html
## append HTML content to PowerPoint slides
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Add a shape
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle:wasmModule.RectangleF.FromLTRB(150, 100, 350, 300)});

// Clear default paragraphs
shape.TextFrame.Paragraphs.Clear();

let code = "<html><body><p>This is a paragraph</p></body></html>";

// Append HTML and generate a paragraph with default style in PPT document
shape.TextFrame.Paragraphs.AddFromHtml(code);

let codeColor = "<html><body><p style=\" color:black \">This is a paragraph</p></body></html>";

// Append HTML with black setting
shape.TextFrame.Paragraphs.AddFromHtml(codeColor);

// Add another shape
let shape1 = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle:wasmModule.RectangleF.FromLTRB(350, 100, 550, 300)});

// Clear default paragraph
shape1.TextFrame.Paragraphs.Clear();

// Change the fill format of shape
shape1.Fill.FillType = wasmModule.FillFormatType.Solid;
shape1.Fill.SolidColor.Color = wasmModule.Color.get_White();

// Append HTML
shape1.TextFrame.Paragraphs.AddFromHtml(code);
const par = shape1.TextFrame.Paragraphs.get_Item(0);

// Change the fill color for paragraph
for (let i = 0;i < par.TextRanges.Count;i++){
    par.TextRanges.get_Item(i).Fill.FillType = wasmModule.FillFormatType.Solid;
    par.TextRanges.get_Item(i).Fill.SolidColor.Color = wasmModule.Color.get_Black();
}
```

---

# spire.presentation javascript text shape
## auto fit text or shape in PowerPoint
```javascript
// Set the AutofitType property to Shape
let textShape2 = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle:wasmModule.RectangleF.FromLTRB(150, 100, 300, 180)});

// Add text in the shape
textShape2.TextFrame.Text = "Resize shape to fit text.";
textShape2.TextFrame.AutofitType = wasmModule.TextAutofitType.Shape;

// Set the AutofitType property to Normal
let textShape1 = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle:wasmModule.RectangleF.FromLTRB(400, 100, 550, 180)});
textShape1.TextFrame.Text = "Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape.";
textShape1.TextFrame.AutofitType = wasmModule.TextAutofitType.Normal;
```

---

# Spire.Presentation JavaScript Borders and Shading
## Set borders and shading for shapes in PowerPoint presentations
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Get the first shape from the first slide
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Set line color and width of the border
shape.Line.FillType = wasmModule.FillFormatType.Solid;
shape.Line.Width = 3;
shape.Line.SolidFillColor.Color = wasmModule.Color.get_LightYellow();

// Set the gradient fill color of shape
shape.Fill.FillType = wasmModule.FillFormatType.Gradient;
shape.Fill.Gradient.GradientShape = wasmModule.GradientShapeType.Linear;
shape.Fill.Gradient.GradientStops.Append({position:1,color:wasmModule.Color.get_LightBlue()});
shape.Fill.Gradient.GradientStops.Append({position:0,color:wasmModule.Color.get_LightSkyBlue()});

// Set the shadow for the shape
let shadow = wasmModule.OuterShadowEffect.Create();
shadow.BlurRadius = 20;
shadow.Direction = 30;
shadow.Distance = 8;
shadow.ColorFormat.Color = wasmModule.Color.get_LightSeaGreen();
shape.EffectDag.OuterShadowEffect = shadow;

// Save to file
ppt.SaveToFile({file:outputFileName,fileFormat:wasmModule.FileFormat.Pptx2013});

// Clean up resources
ppt.Dispose();
```

---

# spire.presentation javascript bullets
## add numbered bullets to paragraphs in PowerPoint
```javascript
// Get the second shape from the first slide
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(1);

// Loop through paragraphs in the shape
for (let i = 0; i < shape.TextFrame.Paragraphs.Count; i++) {
  // Add the bullets
  shape.TextFrame.Paragraphs.get_Item(i).BulletType = wasmModule.TextBulletType.Numbered;
  shape.TextFrame.Paragraphs.get_Item(i).BulletStyle = wasmModule.NumberedBulletStyle.BulletRomanLCPeriod;
}
```

---

# Spire.Presentation JavaScript Text Styling
## Change font and color of text in a PowerPoint presentation
```javascript
// Get the first shape from the first slide
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(0);

let paras = shape.TextFrame.Paragraphs;

// Set the style for the text content in the first paragraph
for(let i = 0;i < paras.get_Item(0).TextRanges.Count;i++){
    paras.get_Item(0).TextRanges.get_Item(i).Fill.FillType = wasmModule.FillFormatType.Solid;
    paras.get_Item(0).TextRanges.get_Item(i).Fill.SolidColor.Color = wasmModule.Color.get_ForestGreen();
    paras.get_Item(0).TextRanges.get_Item(i).LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");
    paras.get_Item(0).TextRanges.get_Item(i).FontHeight = 14;
}

// Set the style for the text content in the third paragraph
for(let i = 0;i < paras.get_Item(2).TextRanges.Count;i++){
    paras.get_Item(2).TextRanges.get_Item(i).Fill.FillType = wasmModule.FillFormatType.Solid;
    paras.get_Item(2).TextRanges.get_Item(i).Fill.SolidColor.Color = wasmModule.Color.get_CornflowerBlue();
    paras.get_Item(2).TextRanges.get_Item(i).LatinFont = wasmModule.TextFont.Create("Calibri");
    paras.get_Item(2).TextRanges.get_Item(i).FontHeight = 16;
    paras.get_Item(2).TextRanges.get_Item(i).TextUnderlineType = wasmModule.TextUnderlineType.Dashed;
}
```

---

# Spire.Presentation JavaScript Copy Paragraph
## Copy paragraph from one PowerPoint presentation to another
```javascript
// Load the source file
let ppt1 = wasmModule.Presentation.Create();
ppt1.LoadFromFile(inputFileName1);

// Get the text from the first shape on the first slide
let sourceshp = ppt1.Slides.get_Item(0).Shapes.get_Item(0);
const text = sourceshp.TextFrame.Text;

// Load the target file
let ppt2 = wasmModule.Presentation.Create();
ppt2.LoadFromFile(inputFileName2);

// Get the first shape on the first slide from the target file
let destshp = ppt2.Slides.get_Item(0).Shapes.get_Item(0);

// Add the text to the target file
destshp.TextFrame.Text += "\n\n" + text;
```

---

# Custom Bullet Numbering in Presentation
## Customize bullet numbering for paragraphs in a presentation slide
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Access the first placeholder in the slide and typecasting it as AutoShape
let tf1 = slide.Shapes.get_Item(1).TextFrame;

// Access the first Paragraph and set bullet style
let para = tf1.Paragraphs.get_Item(0);
para.Depth = 0;
para.BulletType = wasmModule.TextBulletType.Numbered;
para.BulletStyle = wasmModule.NumberedBulletStyle.BulletArabicPeriod;
para.BulletNumber = 2;

// Access the second Paragraph and set bullet style
para = tf1.Paragraphs.get_Item(1);
para.Depth = 0;
para.BulletType = wasmModule.TextBulletType.Numbered;
para.BulletStyle = wasmModule.NumberedBulletStyle.BulletArabicPeriod;
para.BulletNumber = 4;

// Access the third Paragraph and set bullet style
para = tf1.Paragraphs.get_Item(2);
para.Depth = 0;
para.BulletType = wasmModule.TextBulletType.Numbered;
para.BulletStyle = wasmModule.NumberedBulletStyle.BulletArabicPeriod;
para.BulletNumber = 6;

// Access the fourth Paragraph and set bullet style
para = tf1.Paragraphs.get_Item(3);
para.Depth = 0;
para.BulletType = wasmModule.TextBulletType.Numbered;
para.BulletStyle = wasmModule.NumberedBulletStyle.BulletArabicPeriod;
para.BulletNumber = 7;
```

---

# Spire Presentation JavaScript Edit Prompt Text
## Edit prompt text in PowerPoint presentation slides
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Iterate through the slide
for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
    let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
    if (shape.Placeholder != null && shape instanceof wasmModule.IAutoShape) {
        let text = "";
        // Set the text of the title
        if (shape.Placeholder.Type == wasmModule.PlaceholderType.CenteredTitle) {
            text = "custom title create by Spire";
        }
        // Set text of the subtitle
        else if (shape.Placeholder.Type == wasmModule.PlaceholderType.Subtitle) {
            text = "custom subtitle create by Spire";
        }
        shape.TextFrame.Text = text;
    }
}
```

---

# spire.presentation javascript text extraction
## extract text from powerpoint slides
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Initialize an empty array to store extracted text
let sb = [];

// Foreach the slide and extract text
for (let i = 0;i < ppt.Slides.Count;i++){
  // Get the current slide
  let slide = ppt.Slides.get_Item(i);
  for (let j = 0;j < slide.Shapes.Count;j++){
    // Get the current shape on the slide
    let shape = slide.Shapes.get_Item(j);
    if(shape instanceof wasmModule.IAutoShape){
      // Get the paragraphs in the text frame
      let tp = shape.TextFrame.Paragraphs;
      for (let k = 0;k < tp.Count;k++){
        // Add the text of each paragraph to the array
        sb.push(tp.get_Item(k).Text + "\n");
      }
    }
  }
}
// Join all extracted text into a single string
let str = sb.join("");
```

---

# spire.presentation javascript textframe
## get textframe effective data from powerpoint
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

//Get the first slide
let slide = ppt.Slides.get_Item(0);

//Get a first shape from the slide
let shape = slide.Shapes.get_Item(0);

// Get the TextFrame from shape
let textFrameFormat = shape.TextFrame;

// Initialize an empty array to store extracted text
let str = [];

// Add the anchoring type of the text frame to the string
str.push("Anchoring type: " + textFrameFormat.AnchoringType + "\n");

// Add the autofit type of the text frame to the string
str.push("Autofit type: " + textFrameFormat.AutofitType + "\n");

// Add the vertical text type of the text frame to the string
str.push("Text vertical type: " + textFrameFormat.VerticalTextType + "\n");

// Add the margins of the text frame to the string
str.push("Margins" + "\n");
str.push("   Left: " + textFrameFormat.MarginLeft + "\n");
str.push("   Top: " + textFrameFormat.MarginTop + "\n");
str.push("   Right: " + textFrameFormat.MarginRight + "\n");
str.push("   Bottom: " + textFrameFormat.MarginBottom + "\n");

// Join all extracted text into a single string
let content = str.join("");
```

---

# Get Text Style Effective Data
## Extract text style information from paragraphs and text ranges in a PowerPoint slide

```javascript
// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Get a first shape from the slide
let shape = slide.Shapes.get_Item(0);

// Initialize an empty array to store extracted text
let str = [];

for (let p = 0; p < shape.TextFrame.Paragraphs.Count; p++)
{
  let paragraph = shape.TextFrame.Paragraphs.get_Item(p);
  str.push("Text style for Paragraph " + p + " :");

  // Get the paragraph style
  str.push(" Indent: " + paragraph.Indent);
  str.push(" Alignment: " + paragraph.Alignment);
  str.push(" Font alignment: " + paragraph.FontAlignment);
  str.push(" Hanging punctuation: " + paragraph.HangingPunctuation);
  str.push(" Line spacing: " + paragraph.LineSpacing);
  str.push(" Space before: " + paragraph.SpaceBefore);
  str.push(" Space after: " + paragraph.SpaceAfter.toString());
  str.push("\r\n");
  
  for (let r = 0; r < paragraph.TextRanges.Count; r++)
  {
      let textRange = paragraph.TextRanges.get_Item(r);
      str.push("  Text style for Paragraph " + p + " TextRange " + r + " :");

      // Get the text range style
      str.push("    Font height: " + textRange.FontHeight);
      str.push("    Language: " + textRange.Language);
      str.push("    Font: " + textRange.LatinFont.FontName);
      str.push("");
  }
}
```

---

# Spire.Presentation JavaScript Text Highlighting
## Highlight specified text in a presentation
```javascript
// Get the specified shape
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(1);

// Create a new instance of TextHighLightingOptions
let options = wasmModule.TextHighLightingOptions.Create();

// Set the option to highlight only whole words
options.WholeWordsOnly = true;

// Set the option to make the text highlighting case sensitive
options.CaseSensitive = true;

// Highlight the text "Spire" in the text frame with yellow color using the specified options
shape.TextFrame.HighLightText("Spire", wasmModule.Color.get_Yellow(), options);
```

---

# Spire.Presentation JavaScript Paragraph Indentation
## Indent paragraphs in a PowerPoint presentation
```javascript
// Get the first shape from the first slide
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Retrieve the paragraphs
let paras = shape.TextFrame.Paragraphs;

// Set the paragraph style for first paragraph
paras.get_Item(0).Indent = 20;
paras.get_Item(0).LeftMargin = 10;
paras.get_Item(0).SpaceAfter = 10;

// Set the paragraph style of the third paragraph
paras.get_Item(2).Indent = -100;
paras.get_Item(2).LeftMargin = 40;
paras.get_Item(2).SpaceBefore = 0;
paras.get_Item(2).SpaceAfter = 0;
```

---

# Spire.Presentation Line Spacing
## Set paragraph line spacing in PowerPoint presentation
```javascript
// Create a new rectangle shape on the first slide
let shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle:wasmModule.RectangleF.FromLTRB(50, 100, ppt.SlideSize.Size.Width - 50, 400)});

// Set the line color of the shape to white
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

// Set the fill type of the shape to none (transparent)
shape.Fill.FillType = wasmModule.FillFormatType.None;

// Clear any existing paragraphs in the text frame of the shape
shape.TextFrame.Paragraphs.Clear();

// Append a new text frame with specified text to the shape
shape.AppendTextFrame("Spire.Presentation for .NET is a professional PowerPointÂ® compatible API that enables developers to"
    + "create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform."
    + "From Spire.Presentation v 3.7.5, Spire.Presentation starts to support .NET Core, .NET standard.");

// Retrieve the text range from the shape's text frame
let textRange = shape.TextFrame.TextRange;

// Set the fill type of the text to solid
textRange.Fill.FillType = wasmModule.FillFormatType.Solid;

// Set the solid color of the text to blue violet
textRange.Fill.SolidColor.Color = wasmModule.Color.get_BlueViolet();

// Set the font height of the text to 20 points
textRange.FontHeight = 20;

// Set the Latin font of the text to "Lucida Sans Unicode"
textRange.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");

// Adjust paragraph spacing for the first paragraph in the text frame
shape.TextFrame.Paragraphs.get_Item(0).SpaceBefore = 100;
shape.TextFrame.Paragraphs.get_Item(0).SpaceAfter = 100;
shape.TextFrame.Paragraphs.get_Item(0).LineSpacing = 150;
```

---

# Mix Font Styles in PowerPoint Presentation
## Apply different font styles to specific text ranges within a paragraph
```javascript
// Get the second shape of the first slide
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(1);

// Get the text from the shape
let originalText = shape.TextFrame.Text;

// Split the string by specified words and return substrings to a string array
let keywords = ["bold", "red", "underlined", "bigger font size"];
let regex = new RegExp(keywords.map(keywords => keywords.replace(/[-\/\\^$*+?.()|[\]{}]/g,'\\$&')).join('|'),'g');
let splitArray = originalText.split(regex).filter(Boolean);

// Remove the paragraph from TextRange
let tp = shape.TextFrame.TextRange.Paragraph;
tp.TextRanges.Clear();

// Append normal text that is in front of 'bold' to the paragraph
let tr = wasmModule.TextRange.Create(splitArray[0]);
tp.TextRanges.Append(tr);

// Set font style of the text 'bold' as bold
tr = wasmModule.TextRange.Create("bold");
tr.IsBold = wasmModule.TriState.True;
tp.TextRanges.Append(tr);

// Append normal text that is in front of 'red' to the paragraph
tr = wasmModule.TextRange.Create(splitArray[1]);
tp.TextRanges.Append(tr);

// Set the color of the text 'red' as red
tr = wasmModule.TextRange.Create("red");
tr.Fill.FillType = wasmModule.FillFormatType.Solid;
tr.Format.Fill.SolidColor.Color = wasmModule.Color.get_Red();
tp.TextRanges.Append(tr);

// Append normal text that is in front of 'underlined' to the paragraph
tr = wasmModule.TextRange.Create(splitArray[2]);
tp.TextRanges.Append(tr);

// Underline the text 'undelined'
tr = wasmModule.TextRange.Create("underlined");
tr.TextUnderlineType = wasmModule.TextUnderlineType.Single;
tp.TextRanges.Append(tr);

// Append normal text that is in front of 'bigger font size' to the paragraph
tr = wasmModule.TextRange.Create(splitArray[3]);
tp.TextRanges.Append(tr);

// Set a large font for the text 'bigger font size'
tr = wasmModule.TextRange.Create("bigger font size");
tr.FontHeight = 35;
tp.TextRanges.Append(tr);

// Append other normal text
tr = wasmModule.TextRange.Create(splitArray[4]);
tp.TextRanges.Append(tr);
```

---

# Spire.Presentation JavaScript Text Style Modification
## Find and modify the style of the first matched text in a PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Find first "Spire"
let text = "Spire";
let textRange = ppt.Slides.get_Item(0).FindFirstTextAsRange(text);

// Modify the style
textRange.Fill.FillType = wasmModule.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = wasmModule.Color.get_Red();
textRange.FontHeight = 28;
textRange.LatinFont = wasmModule.TextFont.Create("Calibri");
textRange.IsBold = wasmModule.TriState.True;
textRange.IsItalic = wasmModule.TriState.True;
textRange.TextUnderlineType = wasmModule.TextUnderlineType.Double;
textRange.TextStrikethroughType = wasmModule.TextStrikethroughType.Single;
```

---

# Spire.Presentation JavaScript Multiple Level Bullets
## Create multiple level bullets in PowerPoint presentation
```javascript
// Retrieve the TextFrame of the second shape
let tf1 = slide.Shapes.get_Item(1).TextFrame;

// Access the first Paragraph and set bullet style
let para = tf1.Paragraphs.get_Item(0);
para.BulletType = wasmModule.TextBulletType.Symbol;
para.BulletChar = 8226;
para.Depth = 0;

// Access the second Paragraph and set bullet style
para = tf1.Paragraphs.get_Item(1);
para.BulletType = wasmModule.TextBulletType.Symbol;
para.BulletChar = 45;
para.Depth = 1;

// Access the third Paragraph and set bullet style
para = tf1.Paragraphs.get_Item(2);
para.BulletType = wasmModule.TextBulletType.Symbol;
para.BulletChar = 8226;
para.Depth = 2;

// Access the fourth Paragraph and set bullet style
para = tf1.Paragraphs.get_Item(3);
para.BulletType = wasmModule.TextBulletType.Symbol;
para.BulletChar = 45;
para.Depth = 3;
```

---

# spire presentation javascript multiple paragraphs
## create and format multiple paragraphs in a PowerPoint presentation
```javascript
// Access the first slide
let slide = ppt.Slides.get_Item(0);

// Add an AutoShape of rectangle type
let rec = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 250, 150, (500 + ppt.SlideSize.Size.Width / 2 - 250), 300);
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({ shapeType: wasmModule.ShapeType.Rectangle, rectangle: rec });

// Access TextFrame of the AutoShape
let tf = shape.TextFrame;

// Create Paragraphs and TextRanges with different text formats
let para0 = tf.Paragraphs.get_Item(0);
let textRange1 = wasmModule.TextRange.Create("");
let textRange2 = wasmModule.TextRange.Create("");
para0.TextRanges.Append(textRange1);
para0.TextRanges.Append(textRange2);

let para1 = wasmModule.TextParagraph.Create();
tf.Paragraphs._Append(para1);
let textRange11 = wasmModule.TextRange.Create("");
let textRange12 = wasmModule.TextRange.Create("");
let textRange13 = wasmModule.TextRange.Create("");
para1.TextRanges.Append(textRange11);
para1.TextRanges.Append(textRange12);
para1.TextRanges.Append(textRange13);

let para2 = wasmModule.TextParagraph.Create();
tf.Paragraphs._Append(para2);
let textRange21 = wasmModule.TextRange.Create("");
let textRange22 = wasmModule.TextRange.Create("");
let textRange23 = wasmModule.TextRange.Create("");
para2.TextRanges.Append(textRange21);
para2.TextRanges.Append(textRange22);
para2.TextRanges.Append(textRange23);

// Iterate through the first three paragraphs
for (let i = 0; i < 3; i++) {
  // Iterate through the first three text ranges in each paragraph
  for (let j = 0; j < 3; j++) {
    // Set the text for each text range
    tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).Text = "TextRange " + j.toString();
    // Apply formatting based on the index of the text range
    if (j == 0) {
      // Format for the first text range
      tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).Fill.FillType = wasmModule.FillFormatType.Solid;
      tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).Fill.SolidColor.Color = wasmModule.Color.get_LightBlue();
      tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).Format.IsBold = wasmModule.TriState.True;
      tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).FontHeight = 15;
    }
    else if (j == 1) {
      // Format for the second text range
      tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).Fill.FillType = wasmModule.FillFormatType.Solid;
      tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).Fill.SolidColor.Color = wasmModule.Color.get_Blue();
      tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).Format.IsItalic = wasmModule.TriState.True;
      tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).FontHeight = 18;
    }
  }
}
```

---

# PowerPoint Picture Custom Bullet Style
## Add picture as custom bullet style in PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Get the second shape on the first slide
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(1);

// Traverse through the paragraphs in the shape
for (let i = 0; i < shape.TextFrame.Paragraphs.Count; i++) {
  let paragraph = shape.TextFrame.Paragraphs.get_Item(i);
  // Set the bullet style of paragraph as picture
  paragraph.BulletType = wasmModule.TextBulletType.Picture;
  // Load a picture
  let stream = wasmModule.Stream.CreateByFile(inputFileImageName);
  // Add the picture as the bullet style of paragraph
  paragraph.BulletPicture.EmbedImage = ppt.Images.Append({ stream: stream });
}
```

---

# spire.presentation javascript text box
## remove text box from powerpoint slide
```javascript
// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Traverse all the shapes in slide
for (let i = slide.Shapes.Count - 1; i >= 0;i--) {
    //Remove all shapes
    let shape = slide.Shapes.get_Item(i);
    slide.Shapes.Remove(shape);
}
```

---

# Spire Presentation JavaScript Text Replacement
## Replace text in PowerPoint presentation slides
```javascript
// Replaces specific tags in the text of shapes within a slide
function ReplaceTags(pSlide, TagValues) {
  for (let i = 0; i < pSlide.Shapes.Count; i++) {
    let curShape = pSlide.Shapes.get_Item(i);
    if (curShape instanceof wasmModule.IAutoShape) {
      for (let j = 0; j < curShape.TextFrame.Paragraphs.Count; j++) {
        let tp = curShape.TextFrame.Paragraphs.get_Item(j);
        for (let [key, value] of TagValues.entries()) {
          let curKey = key;
          let txt = tp.Text;
          tp.Text = findAllIndices(txt, curKey, value);
        }
      }
    }
  }
};

// Find and replace all occurrences of a value in a string
function findAllIndices(str, value, replaceValue) {
  let result = str;
  let index = 0;
  while ((index = result.indexOf(value, index)) !== -1) {
    result = result.substring(0, index) + replaceValue + result.substring(index + value.length);
    index += replaceValue.length;
  }
  return result;
};
```

---

# Spire.Presentation JavaScript Text Replacement
## Replace text while retaining original style in PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Replace the first occurrence of the word "use" with "test"
ppt.Slides.get_Item(0).ReplaceFirstText("use", "test", true);

// Replace all occurrences of the word "Spire" with "new spire"
ppt.Slides.get_Item(1).ReplaceAllText("Spire", "new spire", true);
```

---

# PowerPoint Text Replacement with Regex
## Replace text in PowerPoint presentation using regular expressions
```javascript
// Regex for all words
let regex = wasmModule.Regex.Create("\\d+.\\d+|\\w+");

// New string value
let newvalue = "This is the test!";

// Loop through shapes and replace text
for (let i = 0;i < ppt.Slides.Count;i++){
    let slide = ppt.Slides.get_Item(i);
    for (let j = 0;j < slide.Shapes.Count;j++){
        let shape = slide.Shapes.get_Item(j);
        shape.ReplaceTextWithRegex(regex,newvalue);
    }
}
```

---

# spire.presentation javascript text rotation
## rotate text in presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);       

// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Get the first shapes in the slide
let shape = slide.Shapes.get_Item(0);

// Set the vertical text orientation of the shape's text frame to 270 degrees (vertical)
shape.TextFrame.VerticalTextType = wasmModule.VerticalTextType.Vertical270;
```

---

# Spire.Presentation JavaScript 3D Text Effects
## Set 3D effects for text in PowerPoint presentations
```javascript
// Set 3D effect for text
shape.TextFrame.TextThreeD.ShapeThreeD.PresetMaterial = wasmModule.PresetMaterialType.Matte;
shape.TextFrame.TextThreeD.LightRig.PresetType = wasmModule.PresetLightRigType.Sunrise;
shape.TextFrame.TextThreeD.ShapeThreeD.TopBevel.PresetType = wasmModule.BevelPresetType.Circle;
shape.TextFrame.TextThreeD.ShapeThreeD.ContourColor.Color = wasmModule.Color.get_Green();
shape.TextFrame.TextThreeD.ShapeThreeD.ContourWidth = 3;
```

---

# Spire.Presentation JavaScript Text Frame Anchor
## Set text frame anchor type to bottom alignment
```javascript
// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Get the first shape from the slide
let shape = slide.Shapes.get_Item(0);

// Set the anchoring type of the shape's text frame to align the text at the bottom of the shape
shape.TextFrame.AnchoringType = wasmModule.TextAnchorType.Bottom;
```

---

# spire.presentation javascript text frame
## set columns count of text frame in PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Get the first shape in first slide and set column count of text for it
let shape1 = ppt.Slides.get_Item(0).Shapes.get_Item(0);
shape1.TextFrame.ColumnCount = 3;

// Get the second shape in second slide and set column count of text for it
let shape2 = ppt.Slides.get_Item(1).Shapes.get_Item(0);
shape2.TextFrame.ColumnCount = 2;

// Save to file
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# Setting Custom Fonts in Presentation
## Demonstrates how to set custom fonts in a PowerPoint presentation
```javascript
// Load the ARIALUNI.TTF font file into the virtual file system (VFS)
await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Create a rectangle using the specified coordinates
let rec = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 250, 80, (500 + ppt.SlideSize.Size.Width / 2 - 250), 230);

// Append a rectangle shape to the first slide using the defined rectangle
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:rec});

// Set the line color of the rectangle shape to white
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

// Set the fill type of the rectangle shape to none, making it transparent
shape.Fill.FillType = wasmModule.FillFormatType.None;

// Add text to the shape
shape.AppendTextFrame("Hello World!");

// Set the custom font folder
ppt.SetCustomFontsFolder("/Library/Fonts/");

// Set the font and fill style of the text
let textRange = shape.TextFrame.TextRange;
textRange.Fill.FillType = wasmModule.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
textRange.FontHeight = 66;
textRange.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");
```

---

# Spire Presentation JavaScript Font Settings
## Set paragraph font properties including font family, bold, italic, and color
```javascript
// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Access the first and second placeholder in the slide and typecasting it as AutoShape
let tf1 = slide.Shapes.get_Item(0).TextFrame;
let tf2 = slide.Shapes.get_Item(1).TextFrame;

// Access the first Paragraph
let para1 = tf1.Paragraphs.get_Item(0);
let para2 = tf2.Paragraphs.get_Item(0);

// Justify the paragraph
para2.Alignment = wasmModule.TextAlignmentType.Justify;

// Access the first text range
let textRange1 = para1.FirstTextRange;
let textRange2 = para2.FirstTextRange;

// Define new fonts
let fd1 = wasmModule.TextFont.Create("Elephant");
let fd2 = wasmModule.TextFont.Create("Castellar");

// Assign new fonts to text range
textRange1.LatinFont = fd1;
textRange2.LatinFont = fd2;

// Set font to Bold
textRange1.Format.IsBold = wasmModule.TriState.True;
textRange2.Format.IsBold = wasmModule.TriState.False;

// Set font to Italic
textRange1.Format.IsItalic = wasmModule.TriState.False;
textRange2.Format.IsItalic = wasmModule.TriState.True;

// Set font color
textRange1.Fill.FillType = wasmModule.FillFormatType.Solid;
textRange1.Fill.SolidColor.Color = wasmModule.Color.get_Purple();
textRange2.Fill.FillType = wasmModule.FillFormatType.Solid;
textRange2.Fill.SolidColor.Color = wasmModule.Color.get_Peru();
```

---

# Spire.Presentation JavaScript Right-to-Left Columns
## Set right-to-left column direction in PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Get the second shape
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(1);

// Set columns style to right-to-left
shape.TextFrame.RightToLeftColumns = true;
```

---

# spire presentation javascript text shadow
## set shadow effect for text in powerpoint
```javascript
// Add a new rectangle shape to the first slide
let shape = slide.Shapes.AppendShape({shapeType: wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(120, 100, 570, 300)});
shape.Fill.FillType = wasmModule.FillFormatType.None;

// Add the text to the shape and set the font for the text
shape.AppendTextFrame("Text shading on slides");
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Arial Black");
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).FontHeight = 21;
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_Black();

// Add outer shadow and set all necessary parameters
let Shadow = wasmModule.OuterShadowEffect.Create();

Shadow.BlurRadius = 0;
Shadow.Direction = 50;
Shadow.Distance = 10;
Shadow.ColorFormat.Color = wasmModule.Color.get_LightBlue();

// Apply the shadow effect to the text
shape.TextFrame.TextRange.EffectDag.OuterShadowEffect = Shadow;
```

---

# Spire.Presentation JavaScript Text Direction
## Set text direction in PowerPoint slides
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Append a shape with text to the first slide
let textboxShape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType: wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(250, 70, 350, 470)});
textboxShape.ShapeStyle.LineColor.Color = wasmModule.Color.get_Transparent();
textboxShape.Fill.FillType = wasmModule.FillFormatType.Solid;
textboxShape.Fill.SolidColor.Color = wasmModule.Color.get_LightBlue();
textboxShape.TextFrame.Text = "You Are Welcome Here";

// Set the text direction to vertical
textboxShape.TextFrame.VerticalTextType = wasmModule.VerticalTextType.Vertical;

// Append another shape with text to the slide
textboxShape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType: wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(350, 70, 450, 470)});
textboxShape.ShapeStyle.LineColor.Color = wasmModule.Color.get_Transparent();
textboxShape.Fill.FillType = wasmModule.FillFormatType.Solid;
textboxShape.Fill.SolidColor.Color = wasmModule.Color.get_LightGray();

// Append some asian characters
textboxShape.TextFrame.Text = "欢迎光临";

// Set the VerticalTextType as EastAsianVertical to avoid rotating text 90 degrees
textboxShape.TextFrame.VerticalTextType = wasmModule.VerticalTextType.EastAsianVertical;
```

---

# Spire.Presentation JavaScript Text Font Properties
## Set text font properties in a presentation including font family, bold, italic, underline, size, and color
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Add a new shape to the PPT document
let rec = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 250, 80, (500 + ppt.SlideSize.Size.Width / 2 - 250), 230);
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: rec});

shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
shape.Fill.FillType = wasmModule.FillFormatType.None;

// Add text to the shape
shape.AppendTextFrame("Welcome to use Spire.Presentation");

let textRange = shape.TextFrame.TextRange;

// Set the font
textRange.LatinFont = wasmModule.TextFont.Create("Times New Roman");

// Set bold property of the font
textRange.IsBold = wasmModule.TriState.True;

// Set italic property of the font
textRange.IsItalic = wasmModule.TriState.True;

// Set underline property of the font
textRange.TextUnderlineType = wasmModule.TextUnderlineType.Single;

// Set the height of the font
textRange.FontHeight = 50;

// Set the color of the font
textRange.Fill.FillType = wasmModule.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
```

---

# Spire.Presentation JavaScript Text Margins
## Set text margins for shapes in PowerPoint presentations
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Append a new shape
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(50, 100, 500, 250)});

// Set text for the shape
shape.TextFrame.Text = "Using Spire.Presentation, developers will find an easy and effective method to create, read, write, modify, convert and print PowerPoint files. It's worthwhile for you to try this amazing product.";

// Set the margins for the text frame
shape.TextFrame.MarginTop = 10;
shape.TextFrame.MarginBottom = 35;
shape.TextFrame.MarginLeft = 15;
shape.TextFrame.MarginRight = 30;
```

---

# spire.presentation javascript text transparency
## set text transparency with different alpha values in PowerPoint
```javascript
// Add another shape to the PPT document
let textboxShape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(100, 100, 400, 220)});
textboxShape.ShapeStyle.LineColor.Color = wasmModule.Color.get_Transparent();
textboxShape.Fill.FillType = wasmModule.FillFormatType.None;

// Remove default blank paragraphs
textboxShape.TextFrame.Paragraphs.Clear();

// Add three paragraphs, apply color with different alpha values to text
let alpha = 55;
for (let i = 0; i < 3; i++) {
    textboxShape.TextFrame.Paragraphs._Append(wasmModule.TextParagraph.Create());
    textboxShape.TextFrame.Paragraphs.get_Item(i).TextRanges.Append(wasmModule.TextRange.Create("Text Transparency"));
    textboxShape.TextFrame.Paragraphs.get_Item(i).TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
    textboxShape.TextFrame.Paragraphs.get_Item(i).TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.FromArgb({alpha,baseColor: wasmModule.Color.get_Purple()});
    alpha += 100;
}
```

---

# spire.presentation javascript superscript subscript
## add superscript and subscript text to PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Append a rectangle shape to the slide
let shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(150, 100, 350, 150)});
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
shape.Fill.FillType = wasmModule.FillFormatType.None;
shape.TextFrame.Paragraphs.Clear();

// Append a text frame with the initial text "Test"
shape.AppendTextFrame("Test");

// Append the superscript text to the paragraph
let tr = wasmModule.TextRange.Create("superscript");
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.Append(tr);

// Set the script distance for the superscript text
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(1).Format.ScriptDistance = 30;

// Set the style for the text range
let textRange = shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0);
textRange.Fill.FillType = wasmModule.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = wasmModule.Color.get_Black();
textRange.FontHeight = 20;
textRange.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");

textRange = shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(1);
textRange.Fill.FillType = wasmModule.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
textRange.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");

// Append another rectangle shape to the slide
shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(150, 150, 350, 200)});
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
shape.Fill.FillType = wasmModule.FillFormatType.None;
shape.TextFrame.Paragraphs.Clear();

// Append a text frame with the initial text "Test" again
shape.AppendTextFrame("Test");

// Append the subscript text to the paragraph
tr = wasmModule.TextRange.Create("subscript");
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.Append(tr);

// Set the script distance for the subscript text
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(1).Format.ScriptDistance = -25;

// Set the style for the text range
textRange = shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0);
textRange.Fill.FillType = wasmModule.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = wasmModule.Color.get_Black();
textRange.FontHeight = 20;
textRange.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");

textRange = shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(1);
textRange.Fill.FillType = wasmModule.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
textRange.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");
```

---

# Spire.Presentation JavaScript Master Slide
## Add image to master slide in PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Get the master
let master = ppt.Masters.get_Item(0);

// Append image to slide master
let rff = wasmModule.RectangleF.FromLTRB(40, 40, 130, 130);
let pic = master.Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: imageName, rectangle: rff });
pic.Line.FillFormat.FillType = wasmModule.FillFormatType.None;

// Add new slide to presentation
ppt.Slides.Append();

// Save to file
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Slide Management
## Add slides to presentation using master layouts
```javascript
// Get Master layouts
let iLayout = ppt.Masters.get_Item(0).Layouts.get_Item(0);

// Append new slide
ppt.Slides.Append({ layout: iLayout });

// Insert new slide
ppt.Slides.Insert({ index: 1, layout: iLayout });
```

---

# spire.presentation javascript slide
## append slide with master layout
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Get the master
let master = ppt.Masters.get_Item(0);

// Get master layout slides
let masterLayouts = master.Layouts;
let layoutSlide = masterLayouts.get_Item(1);

// Append a rectangle to the layout slide
let shape = layoutSlide.Shapes.AppendShape({ shapeType: wasmModule.ShapeType.Rectangle, rectangle: spirepresentation.RectangleF.FromLTRB(10, 50, 110, 130) });

// Add a text into the shape and set the style
shape.Fill.FillType = wasmModule.FillFormatType.None;
shape.AppendTextFrame("Layout slide 1");
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Arial Black");
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();

// Append new slide with master layout
ppt.Slides.Append({ slide: ppt.Slides.get_Item(0), layout: master.Layouts.get_Item(1) });

// Another way to append new slide with master layout
ppt.Slides.Insert({ index: 2, slide: ppt.Slides.get_Item(1), layout: master.Layouts.get_Item(1) });
```

---

# Spire.Presentation JavaScript Slide Master
## Apply slide master settings to a PowerPoint presentation
```javascript
// Get the first slide master from the presentation
let masterSlide = ppt.Masters.get_Item(0);

// Customize the background of the slide master
let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
masterSlide.SlideBackground.Fill.FillType = wasmModule.FillFormatType.Picture;
let image = masterSlide.Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: backgroundPicName, rectangle: rect });
masterSlide.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image.PictureFill.Picture.EmbedImage;

// Change the color scheme
masterSlide.Theme.ColorScheme.Accent1.Color = wasmModule.Color.get_Red();
masterSlide.Theme.ColorScheme.Accent2.Color = wasmModule.Color.get_RosyBrown();
masterSlide.Theme.ColorScheme.Accent3.Color = wasmModule.Color.get_Ivory();
masterSlide.Theme.ColorScheme.Accent4.Color = wasmModule.Color.get_Lavender();
masterSlide.Theme.ColorScheme.Accent5.Color = wasmModule.Color.get_Black();
```

---

# spire.presentation javascript slide manipulation
## change slide position in PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Move the first slide to the second slide position
let slide = ppt.Slides.get_Item(0);
slide.SlideNumber = 2;

// Save to file
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });
```

---

# Spire.Presentation JavaScript Slide Cloning
## Clone slides from one PowerPoint presentation to another
```javascript
// Load PPT document from the specified input file
let ppt1 = wasmModule.Presentation.Create();
ppt1.LoadFromFile(inputFileName1);

// Load PPT document from the specified input file
let ppt2 = wasmModule.Presentation.Create();
ppt2.LoadFromFile(inputFileName2);

// Loop through all slides of source document
for (let i = 0; i < ppt1.Slides.Count; i++) {
  // Append the slide at the end of destination document
  let slide = ppt1.Slides.get_Item(i);
  ppt2.Slides.Append({ slide: slide });
}
```

---

# Clone PowerPoint Master Slides
## Clone master slides from one presentation to another
```javascript
// Load source document from the specified input file
let ppt1 = wasmModule.Presentation.Create();
ppt1.LoadFromFile(inputFileName1);

// Load destination document from the specified input file
let ppt2 = wasmModule.Presentation.Create();
ppt2.LoadFromFile(inputFileName2);

// Add masters from PPT1 to PPT2
for (let i = 0; i < ppt1.Masters.Count; i++) {
  let masterSlide = ppt1.Masters.get_Item(i);
  ppt2.Masters.AppendSlide(masterSlide);
}

// Save to file
ppt2.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt2.Dispose();
```

---

# spire.presentation javascript slide
## clone slide at the end of presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Append the slide at the end of the document
ppt.Slides.Append({ slide: slide });

// Save to file
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# PowerPoint Slide Cloning
## Clone a slide from one presentation to another
```javascript
// Load PPT document from the specified input file
let ppt1 = wasmModule.Presentation.Create();
ppt1.LoadFromFile(inputFileName1);

// Load PPT document from the specified input file
let ppt2 = wasmModule.Presentation.Create();
ppt2.LoadFromFile(inputFileName2);

// Get the first slide
let slide1 = ppt1.Slides.get_Item(0);

// Insert the slide to the specified index in the source presentation
let index = 1;
ppt2.Slides.Insert({index:index, slide:slide1});
```

---

# Spire.Presentation JavaScript Slide Cloning
## Clone a slide within a PowerPoint presentation
```javascript
// Get a list of slides and choose the first slide to be cloned
let slide = ppt.Slides.get_Item(0);

// Insert the desired slide to the specified index in the same presentation
let index = 1;
ppt.Slides.Insert({ index: index, slide: slide });
```

---

# Spire.Presentation JavaScript Slide Creation
## Create presentation slides with shapes, text, and background images
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Add new slide
ppt.Slides.Append();

// Set the background image
for (let i = 0; i < 2; i++) {
  let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
  ppt.Slides.get_Item(i).Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: "background.png", rectangle: rect });
  ppt.Slides.get_Item(i).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();
}

// Add title
let rec_title = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 200, 70, (400 + ppt.SlideSize.Size.Width / 2 - 200), 120);
let shape_title = ppt.Slides.get_Item(0).Shapes.AppendShape({ shapeType: wasmModule.ShapeType.Rectangle, rectangle: rec_title });
shape_title.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
shape_title.Fill.FillType = wasmModule.FillFormatType.None;
let para_title = wasmModule.TextParagraph.Create();
para_title.Text = "E-iceblue";
para_title.Alignment = wasmModule.TextAlignmentType.Center;
para_title.TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Myriad Pro Light");
para_title.TextRanges.get_Item(0).FontHeight = 36;
para_title.TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
para_title.TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_Black();
shape_title.TextFrame.Paragraphs._Append(para_title);

// Append new shape
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({ shapeType: wasmModule.ShapeType.Rectangle, rectangle: wasmModule.RectangleF.FromLTRB(50, 150, 650, 430) });
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
shape.Fill.FillType = wasmModule.FillFormatType.None;
shape.Line.FillType = wasmModule.FillFormatType.None;
// Add text to shape
shape.AppendTextFrame("Welcome to use Spire.Presentation for .NET.");

// Add new paragraph
let pare = wasmModule.TextParagraph.Create();
pare.Text = "";
shape.TextFrame.Paragraphs._Append(pare);

// Add new paragraph
pare = wasmModule.TextParagraph.Create();
pare.Text = "Spire.Presentation for .NET is a professional PowerPoint compatible component that enables developers to create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform. As an independent PowerPoint .NET component, Spire.Presentation for .NET doesn't need Microsoft PowerPoint installed on the machine.";
shape.TextFrame.Paragraphs._Append(pare);

// Set the Font
for (let i = 0; i < shape.TextFrame.Paragraphs.Count; i++) {
  let para = shape.TextFrame.Paragraphs.get_Item(i);
  para.TextRanges.get_Item({ index: 0 }).LatinFont = wasmModule.TextFont.Create("Myriad Pro");
  para.TextRanges.get_Item({ index: 0 }).FontHeight = 24;
  para.TextRanges.get_Item({ index: 0 }).Fill.FillType = wasmModule.FillFormatType.Solid;
  para.TextRanges.get_Item({ index: 0 }).Fill.SolidColor.Color = wasmModule.Color.get_Black();
  para.Alignment = wasmModule.TextAlignmentType.Left;
}

// Append new shape - SixPointedStar
shape = ppt.Slides.get_Item(1).Shapes.AppendShape({ shapeType: wasmModule.ShapeType.SixPointedStar, rectangle: wasmModule.RectangleF.FromLTRB(100, 100, 200, 200) });
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_Orange();
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

// Append new shape
shape = ppt.Slides.get_Item(1).Shapes.AppendShape({ shapeType: wasmModule.ShapeType.Rectangle, rectangle: wasmModule.RectangleF.FromLTRB(50, 250, 650, 300) });
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
shape.Fill.FillType = wasmModule.FillFormatType.None;

// Add text to shape
shape.AppendTextFrame("This is newly added Slide.");

// Set the Font
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Myriad Pro");
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).FontHeight = 24;
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_Black();
shape.TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Left;
shape.TextFrame.Paragraphs.get_Item(0).Indent = 35;
```

---

# spire.presentation javascript slide master
## create slide masters and apply them to slides
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

ppt.SlideSize.Type = wasmModule.SlideSizeType.Screen16x9;

// Add slides
for (let i = 0; i < 4; i++) {
  ppt.Slides.Append();
}

// Get the first default slide master
let first_master = ppt.Masters.get_Item(0);

// Append another slide master
ppt.Masters.AppendSlide(first_master);
let second_master = ppt.Masters.get_Item(1);

// The first slide masters
let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
first_master.SlideBackground.Fill.FillType = wasmModule.FillFormatType.Picture;
let image1 = first_master.Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: pic1Name, rectangle: rect });
first_master.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image1.PictureFill.Picture.EmbedImage;

// The second slide master
second_master.SlideBackground.Fill.FillType = wasmModule.FillFormatType.Picture;
let image2 = second_master.Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: pic2Name, rectangle: rect });
second_master.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image2.PictureFill.Picture.EmbedImage;

// Apply the first master with layout to the first slide
ppt.Slides.get_Item(0).Layout = first_master.Layouts.get_Item(1);

// Apply the second master with layout to other slides
for (let i = 1; i < ppt.Slides.Count; i++) {
  ppt.Slides.get_Item(i).Layout = second_master.Layouts.get_Item(8);
}
```

---

# Detect Used Themes in PowerPoint Presentation
## This code demonstrates how to detect the themes used in each slide of a PowerPoint presentation.

```javascript
// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Initialize an empty array to store theme names
let sb = [];

// Get the theme name of each slide in the document
for (let i = 0; i < ppt.Slides.Count; i++) {
  let themeName = ppt.Slides.get_Item(i).Theme.Name;
  sb.push(themeName);
}

// Join the theme names into a string
let str = sb.join("\r\n");
```

---

# spire.presentation javascript slides
## get slide by index or ID
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

//Get slide by index 0
let slide1 = ppt.Slides.get_Item(0);

// Append a shape in the slide
let shape1 = slide1.Shapes.AppendShape({ shapeType: wasmModule.ShapeType.Rectangle, rectangle: wasmModule.RectangleF.FromLTRB(100, 100, 300, 200) });

// Add text in the shape
shape1.TextFrame.Text = "Get slide by index";

// Get slide by slide ID
let slide2 = ppt.FindSlide(Number(ppt.Slides.get_Item(1).SlideID));

// Append a shape in the slide
let shape2 = slide2.Shapes.AppendShape({ shapeType: wasmModule.ShapeType.Rectangle, rectangle: wasmModule.RectangleF.FromLTRB(100, 100, 300, 200) });

// Add text in the shape
shape2.TextFrame.Text = "Get slide by slide id";
```

---

# PowerPoint Text Extraction
## Extract text from all slides in a PowerPoint presentation
```javascript
// Initialize an empty array to store extracted text
let str = [];

// Foreach the slide and get text
for (let i = 0; i < ppt.Slides.Count; i++) {
  let arrayList = ppt.Slides.get_Item(i).GetAllTextFrame();
  str.push(arrayList);
}

let content = str.join("\r\n");
```

---

# Spire.Presentation JavaScript Hide Slide
## Hide a specific slide in a PowerPoint presentation
```javascript
// Load PPT document
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

// Hide the second slide
ppt.Slides.get_Item(1).Hidden = true;

// Clean up resources
ppt.Dispose();
```

---

# Merge PowerPoint Slides
## Merge selected slides from multiple presentations into a single presentation

```javascript
// Create an instance of presentation document
let ppt = wasmModule.Presentation.Create();

// Remove the first slide
ppt.Slides.RemoveAt(0);

// Load PPT document from the specified input file
let ppt1 = wasmModule.Presentation.Create();
ppt1.LoadFromFile(inputFileName1);

// Load PPT document from the specified input file
let ppt2 = wasmModule.Presentation.Create();
ppt2.LoadFromFile(inputFileName2);

// Append all slides in ppt1 to ppt
for (let i = 0; i < ppt1.Slides.Count; i++) {
  ppt.Slides.Append({ slide: ppt1.Slides.get_Item(i) });
}

// Append the second slide in ppt2 to ppt
ppt.Slides.Append({ slide: ppt2.Slides.get_Item(1) });

// Save to file
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Slide Removal
## Demonstrates how to remove slides by index and reference in a PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Remove slide by index
ppt.Slides.RemoveAt(0);

// Remove slide by its reference
let slide = ppt.Slides.get_Item(1);
ppt.Slides.Remove(slide);

// Save to file
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# PowerPoint Master Layout Cleanup
## Remove unused master layouts from PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Create an array list
let list = [];
for (let i = 0; i < ppt.Slides.Count; i++) {
  // Get the layout used by slide
  let layout = ppt.Slides.get_Item(i).Layout;
  list.push(Number(layout.SlideID));
}

// Loop through masters and layouts
for (let i = 0; i < ppt.Masters.Count; i++) {
  let masterlayouts = ppt.Masters.get_Item(i).Layouts;
  for (let j = masterlayouts.Count - 1; j >= 0; j--) {
    if (!list.includes(Number(masterlayouts.get_Item(j).SlideID))) {
      // Remove unused layout
      masterlayouts.RemoveMasterLayout(j);
    }
  }
}

// Save to file
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Slide Numbering
## Set starting number for slides in PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Set 5 as the starting number
ppt.FirstSlideNumber = 5;

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Slide Title
## Get and set slide titles in a presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Get the title of the first slide
let slideTitle = slide.Title;

// Set the title of the second slide
ppt.Slides.get_Item(1).Title = "Second Slide";
```

---

# Spire.Presentation JavaScript Slides
## Add slides with different layouts to a presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Remove the default slide
ppt.Slides.RemoveAt(0);

// Loop through slide layouts
for (let type in wasmModule.SlideLayoutType){
    //Append slide by specifing slide layout
    ppt.Slides.Append({template:wasmModule.SlideLayoutType[type]});
}
```

---

# spire presentation javascript slide layout
## change slide layout in presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName); 

// Change the layout of slide
ppt.Slides.get_Item(1).Layout = ppt.Masters.get_Item(0).Layouts.get_Item(4);

// Save to file
ppt.SaveToFile({file:outputFileName,fileFormat:wasmModule.FileFormat.Pptx2013});

// Clean up resources
ppt.Dispose();
```

---

# Get Slide Layout Name
## Extracts layout names from PowerPoint slides
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName); 

let builder = [];

// Loop through the slides of PPT document
for (let i = 0; i < ppt.Slides.Count; i++) {
    // Get the name of slide layout
    let name = ppt.Slides.get_Item(i).Layout.Name;
    builder.push(`The name of slide ${i} layout is: ${name}`);
}
let str = builder.join("\r\n");

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Slide Layout
## Set slide layout in PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Remove the first slide
ppt.Slides.RemoveAt(0);

//Append a slide and set the layout for slide
let slide = ppt.Slides.Append({ template: wasmModule.SlideLayoutType.Title });

//Add content for Title and Text
let shape = slide.Shapes.get_Item(0);
shape.TextFrame.Text = "Hello Wolrd! -> This is title";

shape = slide.Shapes.get_Item(1);
shape.TextFrame.Text = "E-iceblue Support Team -> This is content";
```

---

# Spire.Presentation JavaScript Slide Transitions
## Set better transitions for presentation slides
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName); 

// Set the first slide transition as circle
ppt.Slides.get_Item(0).SlideShowTransition.Type = wasmModule.TransitionType.Circle;

// Set the transition time of 3 seconds
ppt.Slides.get_Item(0).SlideShowTransition.AdvanceOnClick = true;
ppt.Slides.get_Item(0).SlideShowTransition.AdvanceAfterTime = 3000;

//Set the second slide transition as comb and set the speed
ppt.Slides.get_Item(1).SlideShowTransition.Type = wasmModule.TransitionType.Comb;
ppt.Slides.get_Item(1).SlideShowTransition.Speed = wasmModule.TransitionSpeed.Slow;

// Set the transition time of 5 seconds
ppt.Slides.get_Item(1).SlideShowTransition.AdvanceOnClick = true;
ppt.Slides.get_Item(1).SlideShowTransition.AdvanceAfterTime = 5000;

// Set the third slide transition as zoom
ppt.Slides.get_Item(2).SlideShowTransition.Type = wasmModule.TransitionType.Zoom;

// Set the transition time of 7 seconds
ppt.Slides.get_Item(2).SlideShowTransition.AdvanceOnClick = true;
ppt.Slides.get_Item(2).SlideShowTransition.AdvanceAfterTime = 7000;
```

---

# Spire.Presentation JavaScript Slides
## Set slide transition advance time
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Traverse all slides in the presentation
for (let i = 0; i < ppt.Slides.Count; i++) {
  // Enable advancing the slide on mouse click
  ppt.Slides.get_Item(i).SlideShowTransition.AdvanceOnClick = true;

  // Set the automatic advance time for the slide
  ppt.Slides.get_Item(i).SlideShowTransition.AdvanceAfterTime = 5000;
}
```

---

# spire.presentation javascript transition effects
## set slide transition effects
```javascript
// Set the transition type for the first slide to a "Cut" transition
ppt.Slides.get_Item(0).SlideShowTransition.Type = wasmModule.TransitionType.Cut;

// Configure the transition to start from a black screen
ppt.Slides.get_Item(0).SlideShowTransition.Value.FromBlack = true;
```

---

# Spire.Presentation JavaScript Transitions
## Set slide transitions in a presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT document from the specified input file
ppt.LoadFromFile(inputFileName);

// Set the first slide transition as push and sound mode
ppt.Slides.get_Item(0).SlideShowTransition.Type = wasmModule.TransitionType.Push;
ppt.Slides.get_Item(0).SlideShowTransition.SoundMode = wasmModule.TransitionSoundMode.StartSound;

// Set the second slide transition as circle and set the speed
ppt.Slides.get_Item(1).SlideShowTransition.Type = wasmModule.TransitionType.Fade;
ppt.Slides.get_Item(1).SlideShowTransition.Speed = wasmModule.TransitionSpeed.Slow;
```

---

# Adding Line to PowerPoint Slide
## Demonstrates how to add a line shape to a slide in a PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Add a line in the slide
let line = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Line, rectangle:wasmModule.RectangleF.FromLTRB(50, 100, 350, 100)});

// Set color of the line
line.ShapeStyle.LineColor.Color = wasmModule.Color.get_Red();
```

---

# Spire.Presentation JavaScript Shapes
## Add lines with arrows to PowerPoint slides
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Add a line to the slides and set its color to red
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Line, rectangle:wasmModule.RectangleF.FromLTRB(150, 100, 250, 200)});
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_Red();
// Set the line end type as StealthArrow
shape.Line.LineEndType = wasmModule.LineEndType.StealthArrow;

// Add a line to the slides and use default color
shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Line, rectangle:wasmModule.RectangleF.FromLTRB(300, 150, 400, 250)});
shape.Rotation = -45;
// Set the line end type as TriangleArrowHead
shape.Line.LineEndType = wasmModule.LineEndType.TriangleArrowHead;

// Add a line to the slides and set its color to Green
shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Line,rectangle: wasmModule.RectangleF.FromLTRB(450, 100, 550, 200)});
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_Green();
shape.Rotation = 90;
// Set the line begin type as TriangleArrowHead
shape.Line.LineBeginType = wasmModule.LineEndType.StealthArrow;
```

---

# Spire.Presentation JavaScript Line Shapes
## Add lines with two points to PowerPoint slides
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Get the first slide
let slide = ppt.Slides.get_Item(0);

//Add line with two points
let line = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Line,start: wasmModule.PointF.Create(50, 50),end: wasmModule.PointF.Create(150, 150)});
line.ShapeStyle.LineColor.Color = wasmModule.Color.get_Red();
line = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Line,start: wasmModule.PointF.Create(150, 150),end: wasmModule.PointF.Create(250, 50)});
line.ShapeStyle.LineColor.Color = wasmModule.Color.get_Blue();
```

---

# spire.presentation javascript round corner rectangle
## add round corner rectangle to PowerPoint slide
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Append a round corner rectangle and set its radius
let shape = ppt.Slides.get_Item(0).Shapes.AppendRoundRectangle(300, 90, 100, 200, 80);
//Set the color and fill style of shape
shape.Fill.FillType =  wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color =  wasmModule.Color.get_LightBlue();
shape.ShapeStyle.LineColor.Color =  wasmModule.Color.get_SkyBlue();
//Rotate the shape to 90 degree
shape.Rotation = 90;
```

---

# Spire.Presentation JavaScript Shapes
## Add various shapes to a PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle, fileName:ImageFileName, rectangle:rect});
ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

// Append new shape - Triangle and set style
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Triangle, rectangle:wasmModule.RectangleF.FromLTRB(115, 130, 215, 230)});
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_LightGreen();
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

// Append new shape - Ellipse
shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Ellipse,rectangle: wasmModule.RectangleF.FromLTRB(290, 130, 440, 230)});
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_LightSkyBlue();
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

// Append new shape - Heart
shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Heart, rectangle:wasmModule.RectangleF.FromLTRB(470, 130, 600, 230)});
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_Red();
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_LightGray();

// Append new shape - FivePointedStar
shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.FivePointedStar, rectangle:wasmModule.RectangleF.FromLTRB(90, 270, 240, 420)});
shape.Fill.FillType = wasmModule.FillFormatType.Gradient;
shape.Fill.SolidColor.Color = wasmModule.Color.get_Black();
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

// Append new shape - Rectangle
shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(320, 290, 420, 410)});
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_Pink();
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_LightGray();

// Append new shape - BentUpArrow
shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.BentUpArrow, rectangle:wasmModule.RectangleF.FromLTRB(470, 300, 620, 400)});

// Set the color of shape
shape.Fill.FillType = wasmModule.FillFormatType.Gradient;
shape.Fill.Gradient.GradientStops.Append({position:1,knownColor: wasmModule.KnownColors.Olive});
shape.Fill.Gradient.GradientStops.Append({position:0,knownColor: wasmModule.KnownColors.PowderBlue});
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
```

---

# Spire.Presentation JavaScript Shape Arrangement
## Arrange shapes in PowerPoint document by bringing forward
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load file
ppt.LoadFromFile(inputFileName);

// Get the specified shape
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Bring the shape forward through SetShapeArrange method
shape.SetShapeArrange(wasmModule.ShapeArrange.BringForward);

// Save the document to the specified path
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Background
## Set background image for PowerPoint slide
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

const ImageFileName = "backgroundImg.png";
let rect =  wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
ppt.Slides.get_Item(0).Shapes.AppendEmbedImage( {shapeType:wasmModule.ShapeType.Rectangle, fileName:ImageFileName, rectangle:rect});

//Add title
let rec_title = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 200, 70, (380 + ppt.SlideSize.Size.Width / 2 - 200), 120);
let shape_title = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: rec_title});
shape_title.Line.FillType = wasmModule.FillFormatType.None;
shape_title.Fill.FillType = wasmModule.FillFormatType.None;
let para_title = wasmModule.TextParagraph.Create();
para_title.Text = "Background Sample";
para_title.Alignment = wasmModule.TextAlignmentType.Center;
para_title.TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");
para_title.TextRanges.get_Item(0).FontHeight = 36;
para_title.TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
para_title.TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_DarkSlateBlue();
shape_title.TextFrame.Paragraphs._Append(para_title);

//Add new shape to PPT document
let rec = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 300, 155, (600 + ppt.SlideSize.Size.Width / 2 - 300), 355);
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:rec});
shape.Line.FillType = wasmModule.FillFormatType.None;
shape.Fill.FillType = wasmModule.FillFormatType.None;

let para = wasmModule.TextParagraph.Create();
para.Text = "Spire.Presentation supports PPT, PPS, PPTX and PPSX presentation formats. It provides functions such as managing text, image, shapes, tables, animations, audio and video on slides. It also support exporting presentation slides to EMF, JPG, TIFF, PDF format etc.";

para.TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
para.TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
para.TextRanges.get_Item(0).FontHeight = 26;
shape.TextFrame.Paragraphs._Append(para);
```

---

# Spire.Presentation JavaScript Shape Copying
## Copy shapes between slides in a PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

ppt.LoadFromFile(inputFileName);

// Define the source slide and target slide
let sourceSlide = ppt.Slides.get_Item(0);
let targetSlide = ppt.Slides.get_Item(1);

// Copy the first shape from the source slide to the target slide
targetSlide.Shapes.AddShape(sourceSlide.Shapes.get_Item(0));

// Save the document to the specified path
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });
```

---

# spire presentation javascript gradient fill
## fill shape with gradient in PowerPoint document
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

ppt.LoadFromFile(inputFileName);

//Get the first shape and set the style to be Gradient
let GradientShape = ppt.Slides.get_Item(0).Shapes.get_Item(0);
GradientShape.Fill.FillType = wasmModule.FillFormatType.Gradient;
GradientShape.Fill.Gradient.GradientStops.Append({position:0, color:wasmModule.Color.get_LightSkyBlue()});
GradientShape.Fill.Gradient.GradientStops.Append({position:1, color:wasmModule.Color.get_LightGray()});
```

---

# spire.presentation javascript pattern fill
## fill shape with pattern in PowerPoint
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Add a rectangle
let rect = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 50, 100, (100 + ppt.SlideSize.Size.Width / 2 - 50), 200);
let shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:rect});

// Set the pattern fill format
shape.Fill.FillType = wasmModule.FillFormatType.Pattern;
shape.Fill.Pattern.PatternType = wasmModule.PatternFillType.Trellis;
shape.Fill.Pattern.BackgroundColor.Color = wasmModule.Color.get_DarkGray();
shape.Fill.Pattern.ForegroundColor.Color = wasmModule.Color.get_Yellow();

// Set the fill format of line
shape.Line.FillType = wasmModule.FillFormatType.Solid;
shape.Line.SolidFillColor.Color = wasmModule.Color.get_Transparent();
```

---

# Fill Shape with Picture in PowerPoint
## Demonstrates how to fill a shape in a PowerPoint slide with a picture using Spire.Presentation for JavaScript
```javascript
//Get the first shape and set the style to be Gradient
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Fill the shape with picture
shape.Fill.FillType = wasmModule.FillFormatType.Picture;
shape.Fill.PictureFill.Picture.Url = picUrlName;
shape.Fill.PictureFill.FillType = wasmModule.PictureFillType.Stretch;
```

---

# PowerPoint Shape Solid Color Fill
## Fill a shape with solid color in a PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Add a rectangle
let rect = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 50, 100, (100 + ppt.SlideSize.Size.Width / 2 - 50), 200);
let shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:rect});

// Fill shape with solid color
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_Yellow();

// Set the fill format of line
shape.Line.FillType = wasmModule.FillFormatType.Solid;
shape.Line.SolidFillColor.Color = wasmModule.Color.get_Gray();
```

---

# Find Shape by Alternative Text
## Function to find a shape in a PowerPoint slide by its alternative text
```javascript
// Find shape in the slide
let shape = FindShape(slide, "Shape1");

function FindShape(slide, altText) {
    // Loop through shapes in the slide
    for (let i = 0; i < slide.Shapes.Count; i++) {
        let shape = slide.Shapes.get_Item(i);
        // Find the shape whose alternative text is altText
        if (shape.AlternativeText == altText) {
            return shape;
        }
    }
    return null;
}
```

---

# spire.presentation javascript get titles
## extract all titles from slides in a PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load file
ppt.LoadFromFile(inputFileName);

// Instantiate a list of IShape objects
let shapelist = [];
// Loop through all slides and all shapes on each slide
for (let i = 0; i < ppt.Slides.Count; i++) {
    let slide = ppt.Slides.get_Item(i);
    for (let j = 0; j < slide.Shapes.Count; j++) {
        let shape = slide.Shapes.get_Item(j);
        if (shape.Placeholder != null) {
            // Get all titles
            switch (shape.Placeholder.Type) {
                case wasmModule.PlaceholderType.Title:
                    shapelist.push(shape);
                    break;
                case wasmModule.PlaceholderType.CenteredTitle:
                    shapelist.push(shape);
                    break;
                case wasmModule.PlaceholderType.Subtitle:
                    shapelist.push(shape);
                    break;
            }
        }
    }
}

// Loop through the list and get the inner text of all shapes in the list
let stringBuilder = [];
stringBuilder.push("Below are all the obtained titles:");
for (let i = 0; i < shapelist.length; i++) {
    let shape1 = shapelist[i];
    stringBuilder.push(shape1.TextFrame.Text);
}
```

---

# Spire.Presentation JavaScript Shape Group
## Get alternative text from shapes in shape groups
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load document from disk
ppt.LoadFromFile(inputFileName);

let builder = [];

//Loop through slides and shapes
for (let i = 0;i < ppt.Slides.Count;i++) {
    let slide = ppt.Slides.get_Item(i);
    for (let j = 0; j < slide.Shapes.Count; j++) {
        let shape = slide.Shapes.get_Item(j);
        if(shape instanceof wasmModule.GroupShape){
            //Find the shape group
            let groupShape = shape;
            for (let k = 0;k < groupShape.Shapes.Count;k++){
                let gShape = groupShape.Shapes.get_Item(k);
                //Append the alternative text in builder
                builder.push(gShape.AlternativeText);
            }
        }
    }
}
```

---

# spire.presentation javascript shapes
## get shapes by placeholder
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

let placeholder = ppt.Slides.get_Item(1).Shapes.get_Item(0).Placeholder;
//Get Shapes by Placeholder
let shapes = ppt.Slides.get_Item(1).GetPlaceholderShapes(placeholder);

let text = "";
//Iterate over all the shapes
for (let i = 0; i < shapes.length; i++){
    //If shape is IAutoShape
    if (shapes[i] instanceof wasmModule.IAutoShape){
        let autoShape = shapes[i];
        if (autoShape.TextFrame != null) {
            text += autoShape.TextFrame.Text + "\r\n";
        }
    }
}
```

---

# Spire.Presentation JavaScript Group Shapes
## This example demonstrates how to create and group shapes in a PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Get the first slide
let slide = ppt.Slides.get_Item(0);

//Create two shapes in the slide
let rectangle = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: wasmModule.RectangleF.FromLTRB(250, 180, 450, 220)});
rectangle.Fill.FillType = wasmModule.FillFormatType.Solid;
rectangle.Fill.SolidColor.KnownColor = wasmModule.KnownColors.SkyBlue;
rectangle.Line.Width = 0.1;
let ribbon = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Ribbon2, rectangle:wasmModule.RectangleF.FromLTRB(290, 155, 410, 235)});
ribbon.Fill.FillType = wasmModule.FillFormatType.Solid;
ribbon.Fill.SolidColor.KnownColor = wasmModule.KnownColors.LightPink;
ribbon.Line.Width = 0.1;

//Add the two shape objects to an array list
let list = [];
list.push(rectangle);
list.push(ribbon);

//Group the shapes in the list
ppt.Slides.get_Item(0).GroupShapes(list);
```

---

# Hide Shape in PowerPoint Presentation
## Find shape by alternative text and hide it
```javascript
// Loop through slides
for (let i = 0; i < ppt.Slides.Count; i++) {
    let slide = ppt.Slides.get_Item(i);
    // Loop through shapes in the slide
    for (let j = 0; j < slide.Shapes.Count; j++) {
        let shape = slide.Shapes.get_Item(j);
        // Find the shape whose alternative text is Shape1
        if (shape.AlternativeText == "Shape1") {
            // Hide the shape
            shape.IsHidden = true;
        }
    }
}
```

---

# spire.presentation javascript shape detection
## detect if shapes are textboxes in a presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load file
ppt.LoadFromFile(inputFileName);

let stringBuilder = [];

for (let i = 0; i < ppt.Slides.Count; i++) {
    let slide = ppt.Slides.get_Item(i);
    for (let j = 0; j < slide.Shapes.Count; j++) {
        let shape = slide.Shapes.get_Item(j);
        if(shape instanceof wasmModule.IAutoShape){
            //Judge if the shape is textbox
            let isTextbox = shape.IsTextBox;
            stringBuilder.push(isTextbox ? "shape is text box" : "shape is not text box")
        }
    }
}
```

---

# spire presentation javascript placeholders
## operate placeholders in powerpoint slides
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load the document from disk
ppt.LoadFromFile(inputFileName);

// Operate placeholders
for (let j = 0; j < ppt.Slides.Count; j++){
    let slide = ppt.Slides.get_Item(j);
    for (let i = 0; i < slide.Shapes.Count; i++) {
        let shape = slide.Shapes.get_Item(i);
        switch (shape.Placeholder.Type) {
            case wasmModule.PlaceholderType.Media:
                shape.InsertVideo(videoFileName);
                break;
            case wasmModule.PlaceholderType.Picture:
                shape.InsertPicture({filepath:imageFileName});
                break;

            case wasmModule.PlaceholderType.Chart:
                shape.InsertChart(wasmModule.ChartType.ColumnClustered);
                break;

            case wasmModule.PlaceholderType.Table:
                shape.InsertTable(3, 2);
                break;

            case wasmModule.PlaceholderType.Diagram:
                shape.InsertSmartArt(wasmModule.SmartArtLayoutType.BasicBlockList);
                break;
        }
    }
}
```

---

# Shape Locking in PowerPoint Presentation
## Demonstrate how to prevent or allow changing shape properties in a PPT document
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Add a rectangle shape to the slide
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(50, 100, 450, 250)});

//Set the shape format
shape.Fill.FillType = wasmModule.FillFormatType.None;
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_LightBlue();
shape.TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Justify;
shape.TextFrame.Text = "Demo for locking shapes:\n    Green/Black stands for editable.\n    Grey stands for non-editable.";
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Arial Rounded MT Bold");
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_Black();

//The changes of selection and rotation are allowed
shape.Locking.RotationProtection = false;
shape.Locking.SelectionProtection = false;
//The changes of size, position, shape type, aspect ratio, text editing and ajust handles are not allowed
shape.Locking.ResizeProtection = true;
shape.Locking.PositionProtection = true;
shape.Locking.ShapeTypeProtection = true;
shape.Locking.AspectRatioProtection = true;
shape.Locking.TextEditingProtection = true;
shape.Locking.AdjustHandlesProtection = true;
```

---

# PowerPoint Shape Removal
## Remove shapes from PowerPoint presentation based on alternative text
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load document from disk
ppt.LoadFromFile(inputFileName);

//Loop through slides
for (let i = 0; i < ppt.Slides.Count; i++){
    let slide = ppt.Slides.get_Item(i);
    //Loop through shapes
    for (let j = 0; j < slide.Shapes.Count; j++){
        let shape = slide.Shapes.get_Item(j);
        //Find the shapes whose alternative text contain "Shape"
        if (shape.AlternativeText.includes("Shape")) {
            slide.Shapes.Remove(shape);
            j--;
        }
    }
}
```

---

# spire presentation javascript shapes
## reorder overlapping shapes in powerpoint
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load file
ppt.LoadFromFile(inputFileName);

//Get the first shape of the first slide
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(0);
//Change the shape's zorder
ppt.Slides.get_Item(0).Shapes.ZOrder({index:1, shape:shape});

// Define the output file name
const outputFileName = "OverlappingShapes_result.pptx";

// Save the document to the specified path
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });
```

---

# Reset Position of Placeholder
## Reset the position of date time and slide number placeholder in a PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the file from disk.
ppt.LoadFromFile(inputFileName);

//Get the first slide from the sample document.
let slide = ppt.Slides.get_Item(0);

for (let i = 0; i < slide.Shapes.Count; i++) {
    let shape = slide.Shapes.get_Item(i);
    //Reset the position of the slide number to the left.
    if(shape.Name.includes("Slide Number Placeholder")){
        shape.Left = 0;
    }else if(shape.Name.includes("Date Placeholder")){
        //Reset the position of the date time to the center.
        shape.Left = ppt.SlideSize.Size.Width / 2;
        //Reset the date time display style.
        shape.TextFrame.TextRange.Paragraph.Text = wasmModule.DateTime.get_Now().ToString({format:"dd.MM.yyyy"});
        shape.TextFrame.IsCentered = true;
    }
}
```

---

# Spire.Presentation JavaScript Shape Manipulation
## Reset shape size and position when changing slide size
```javascript
// Define the original slide size
let currentHeight = ppt.SlideSize.Size.Height;
let currentWidth = ppt.SlideSize.Size.Width;

// Change the slide size as A3
ppt.SlideSize.Type = wasmModule.SlideSizeType.A3;

// Define the new slide size
let newHeight = ppt.SlideSize.Size.Height;
let newWidth = ppt.SlideSize.Size.Width;

// Define the ratio from the old and new slide size
let ratioHeight = newHeight / currentHeight;
let ratioWidth = newWidth / currentWidth;

// Reset the size and position of the shape on the slide
for (let i = 0; i < ppt.Slides.Count; i++) {
    let slide = ppt.Slides.get_Item(i);
    for (let j = 0; j < slide.Shapes.Count; j++) {
        let shape = slide.Shapes.get_Item(j);
        shape.Height = shape.Height * ratioHeight;
        shape.Width = shape.Width * ratioWidth;

        shape.Left = shape.Left * ratioHeight;
        shape.Top = shape.Top * ratioWidth;
    }
}
```

---

# spire.presentation javascript shape rotation
## rotate shapes in a PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

ppt.LoadFromFile(inputFileName);

//Get the shapes
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Set the rotation
shape.Rotation = 60;

ppt.Slides.get_Item(0).Shapes.get_Item(1).Rotation = 120;
ppt.Slides.get_Item(0).Shapes.get_Item(2).Rotation = 180;
ppt.Slides.get_Item(0).Shapes.get_Item(3).Rotation = 240;
```

---

# spire.presentation javascript 3d effects
## set 3D effect for shapes in PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Add shape1 and fill it with color
let shape1 = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.RoundCornerRectangle, rectangle:wasmModule.RectangleF.FromLTRB(150, 150, 300, 300)});
shape1.Fill.FillType = wasmModule.FillFormatType.Solid;
shape1.Fill.SolidColor.KnownColor = wasmModule.KnownColors.SkyBlue;

//Initialize a new instance of the 3-D class for shape1 and set its properties
let effect1 = shape1.ThreeD.ShapeThreeD;
effect1.PresetMaterial = wasmModule.PresetMaterialType.Powder;
effect1.TopBevel.PresetType = wasmModule.BevelPresetType.ArtDeco;
effect1.TopBevel.Height = 4;
effect1.TopBevel.Width = 12;
effect1.BevelColorMode = wasmModule.BevelColorType.Contour;
effect1.ContourColor.KnownColor = wasmModule.KnownColors.LightBlue;
effect1.ContourWidth = 3.5;

//Add shape2 and fill it with color
let shape2 = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Pentagon,rectangle: wasmModule.RectangleF.FromLTRB(400, 150, 550, 300)});
shape2.Fill.FillType = wasmModule.FillFormatType.Solid;
shape2.Fill.SolidColor.KnownColor = wasmModule.KnownColors.LightGreen;

//Initialize a new instance of the 3-D class for shape2 and set its properties
let effect2 = shape2.ThreeD.ShapeThreeD;
effect2.PresetMaterial = wasmModule.PresetMaterialType.SoftEdge;
effect2.TopBevel.PresetType = wasmModule.BevelPresetType.SoftRound;
effect2.TopBevel.Height = 12;
effect2.TopBevel.Width = 12;
effect2.BevelColorMode = wasmModule.BevelColorType.Contour;
effect2.ContourColor.KnownColor = wasmModule.KnownColors.LawnGreen;
effect2.ContourWidth = 5;
```

---

# Spire.Presentation JavaScript Shape Alternative Text
## Set and get alternative text of shapes in a PowerPoint document
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load file
ppt.LoadFromFile(inputFileName);

//Get the first slide
let slide = ppt.Slides.get_Item(0);

//Set the alternative text (title and description)
slide.Shapes.get_Item(0).AlternativeTitle = "Rectangle";
slide.Shapes.get_Item(0).AlternativeText = "This is a Rectangle";

//Get the alternative text (title and description)
let alternativeText = "";
let title = slide.Shapes.get_Item(0).AlternativeTitle;
alternativeText += "Title: " + title + "\r\n";
let description = slide.Shapes.get_Item(0).AlternativeText;
alternativeText += "Description: " + description;
```

---

# Spire.Presentation JavaScript Ellipse Formatting
## Apply formatting to ellipse shape in PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Get the first slide
let slide = ppt.Slides.get_Item(0);

//Add a rectangle
let rect = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 100, 100, (200 + ppt.SlideSize.Size.Width / 2 - 100), 200);
let shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Ellipse, rectangle:rect});

//Set the fill format of shape
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();

//Set the fill format of line
shape.Line.FillType = wasmModule.FillFormatType.Solid;
shape.Line.SolidFillColor.Color = wasmModule.Color.get_DimGray();
```

---

# Spire.Presentation JavaScript Line Formatting
## Set format for lines in PowerPoint shapes
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Set background image
let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName: ImageFileName,rectangle: rect});
ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

//Add a rectangle shape to the slide
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: wasmModule.RectangleF.FromLTRB(100, 150, 300, 250)});
//Set the fill color of the rectangle shape
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_White();
//Apply some formatting on the line of the rectangle
shape.Line.Style = wasmModule.TextLineStyle.ThickThin;
shape.Line.Width = 5;
shape.Line.DashStyle = wasmModule.LineDashStyleType.Dash;
//Set the color of the line of the rectangle
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_SkyBlue();

//Add a ellipse shape to the slide
shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Ellipse,rectangle: wasmModule.RectangleF.FromLTRB(400, 150, 600, 250)});
//Set the fill color of the ellipse shape
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_White();
//Apply some formatting on the line of the ellipse
shape.Line.Style = wasmModule.TextLineStyle.ThickBetweenThin;
shape.Line.Width = 5;
shape.Line.DashStyle = wasmModule.LineDashStyleType.DashDot;
//Set the color of the line of the ellipse
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_OrangeRed();
```

---

# spire.presentation javascript shapes
## set line join styles for shapes in PowerPoint
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Add three shapes
let shape1 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(50, 150, 200, 200)});
let shape2 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: wasmModule.RectangleF.FromLTRB(250, 150, 400, 200)});
let shape3 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(450, 150, 600, 200)});

// Fill shapes
shape1.Fill.FillType = wasmModule.FillFormatType.Solid;
shape1.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
shape2.Fill.FillType = wasmModule.FillFormatType.Solid;
shape2.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
shape3.Fill.FillType = wasmModule.FillFormatType.Solid;
shape3.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();

// Fill lines of shapes
shape1.Line.FillType = wasmModule.FillFormatType.Solid;
shape1.Line.SolidFillColor.Color = wasmModule.Color.get_DarkGray();
shape2.Line.FillType = wasmModule.FillFormatType.Solid;
shape2.Line.SolidFillColor.Color = wasmModule.Color.get_DarkGray();
shape3.Line.FillType = wasmModule.FillFormatType.Solid;
shape3.Line.SolidFillColor.Color = wasmModule.Color.get_DarkGray();

// Set the line width
shape1.Line.Width = 10;
shape2.Line.Width = 10;
shape3.Line.Width = 10;

// Set the join styles of lines
shape1.Line.JoinStyle = wasmModule.LineJoinType.Bevel;
shape2.Line.JoinStyle = wasmModule.LineJoinType.Miter;
shape3.Line.JoinStyle = wasmModule.LineJoinType.Round;

// Add text in shapes
shape1.TextFrame.Text = "Bevel Join Style";
shape2.TextFrame.Text = "Miter Join Style";
shape3.TextFrame.Text = "Round Join Style";
```

---

# Spire.Presentation JavaScript Shape Effects
## Set outline and effects for shapes in PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Get the first slide
let slide = ppt.Slides.get_Item(0);

//Draw a Rectangle shape
let shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(150, 180, 250, 230)});
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_SkyBlue();
//Set outline color
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_Red();
//Set shadow effect
let shadow = wasmModule.PresetShadow.Create();
shadow.ColorFormat.Color = wasmModule.Color.get_LightSkyBlue();
shadow.Preset = wasmModule.PresetShadowValue.FrontRightPerspective;
shadow.Distance = 10.0;
shadow.Direction = 225.0;
shape.EffectDag.PresetShadowEffect = shadow;

//Draw a Ellipse shape
shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Ellipse,rectangle: wasmModule.RectangleF.FromLTRB(400, 150, 500, 250)});
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_SkyBlue();
//Set outline color
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_Yellow();
//Set shadow effect
let glow = wasmModule.GlowEffect.Create();
glow.ColorFormat.Color = wasmModule.Color.get_LightPink();
glow.Radius = 20.0;
shape.EffectDag.GlowEffect = glow;
```

---

# Spire.Presentation JavaScript Shapes
## Set radius for rounded rectangles in a presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Get the first slide
let slide = ppt.Slides.get_Item(0);

//Insert a rectangle with four round corners and set its radius
let shape1 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.RoundCornerRectangle,rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 200, 200)});
shape1.SetRoundRadius(shape1.Width / 3);

//Insert a rectangle with one round corner and set its radius
let shape2 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.OneRoundCornerRectangle,rectangle: wasmModule.RectangleF.FromLTRB(250, 50, 400, 200)});
shape2.SetRoundRadius(shape2.Width / 3);

//Insert a rectangle with one round corner and which one round cornet is snipped and set its radius
let shape3 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.OneSnipOneRoundCornerRectangle,rectangle: wasmModule.RectangleF.FromLTRB(450, 50, 600, 200)});
shape3.SetRoundRadius(shape3.Width / 3);

//Insert a rectangle with two diagonal round corners and set its radius
let shape4 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.TwoDiagonalRoundCornerRectangle,rectangle: wasmModule.RectangleF.FromLTRB(50, 250, 200, 400)});
shape4.SetRoundRadius(shape4.Width / 3);

//Insert a rectangle with two same side round corners and set its radius
let shape5 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.TwoSamesideRoundCornerRectangle, rectangle:wasmModule.RectangleF.FromLTRB(250, 250, 400, 400)});
shape5.SetRoundRadius(shape5.Width / 3);
```

---

# spire presentation rounded rectangle
## set radius of rounded rectangle shapes
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Insert a rounded rectangle and set its radius
ppt.Slides.get_Item(0).Shapes.InsertRoundRectangle(0, 160, 180, 100, 200, 10);

//Append a rounded rectangle and set its radius
let shape = ppt.Slides.get_Item(0).Shapes.AppendRoundRectangle(380, 180, 100, 200, 100);
//Set the color and fill style of shape
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_SeaGreen();
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

//Rotate the shape to 90 degree
shape.Rotation = 90;
```

---

# spire.presentation javascript rectangle formatting
## apply formatting to rectangle shape in powerpoint
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Add a shape
let rect = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 100, 100, (200 + ppt.SlideSize.Size.Width / 2 - 100), 200);
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: rect});

//Set the fill format of shape
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();

//Set the fill format of line
shape.Line.FillType = wasmModule.FillFormatType.Solid;
shape.Line.SolidFillColor.Color = wasmModule.Color.get_DimGray();
```

---

# Spire.Presentation JavaScript Shadow Effect
## Set shadow effect for shapes in PowerPoint presentations
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

let slide = ppt.Slides.get_Item(0);

//Add a shape to slide.
let rect1 = wasmModule.RectangleF.FromLTRB(200, 150, 500, 270);
let shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: rect1});
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_LightBlue();
shape.Line.FillType = wasmModule.FillFormatType.None;
shape.TextFrame.Text = "This demo shows how to apply shadow effect to shape.";
shape.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.TextFrame.TextRange.Fill.SolidColor.Color = wasmModule.Color.get_Black();

//Create an inner shadow effect through InnerShadowEffect object.
let innerShadow = wasmModule.InnerShadowEffect.Create();
innerShadow.BlurRadius = 20;
innerShadow.Direction = 0;
innerShadow.Distance = 0;
innerShadow.ColorFormat.Color = wasmModule.Color.get_Black();

//Apply the shadow effect to shape.
shape.EffectDag.InnerShadowEffect = innerShadow;
```

---

# Spire.Presentation JavaScript Shape Conversion
## Convert shapes in PowerPoint to images
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load the presentation
ppt.LoadFromFile(inputFileName);

// Convert each shape to an image
for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++){
    let images = ppt.Slides.get_Item(0).Shapes.get_Item(i).SaveAsImage();
    let outFileName = `ShapeToImage-${i}.png`;
    images.Save(outFileName);
    
    // Clean up resources
    images.Dispose();
}

// Clean up resources
ppt.Dispose();
```

---

# spire presentation javascript ungroup shapes
## ungroup shapes in a presentation slide
```javascript
let groupShape = ppt.Slides.get_Item(0).Shapes.get_Item(0);
// Ungroup the shapes
ppt.Slides.get_Item(0).Ungroup(groupShape);
```

---

# spire.presentation javascript animation
## add exit animation for shape
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Get the first slide
let slide = ppt.Slides.get_Item(0);

//Add a shape to the slide
let starShape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.FivePointedStar,rectangle: wasmModule.RectangleF.FromLTRB(250, 100, 450, 300)});
starShape.Fill.FillType = wasmModule.FillFormatType.Solid;
starShape.Fill.SolidColor.KnownColor = wasmModule.KnownColors.LightBlue;

//Add random bars effect to the shape
let effect = slide.Timeline.MainSequence.AddEffect(starShape, wasmModule.AnimationEffectType.RandomBars);

//Change effect type from entrance to exit
effect.PresetClassType = wasmModule.TimeNodePresetClassType.Exit;
```

---

# PowerPoint Animation Effects
## Set animations for shapes and slides in PowerPoint presentations
```javascript
//Set the animation of slide to Circle
ppt.Slides.get_Item(0).SlideShowTransition.Type = wasmModule.TransitionType.Circle;

//Append new shape - Triangle
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Triangle,rectangle: wasmModule.RectangleF.FromLTRB(100, 280, 180, 360)});

//Set the color of shape
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

//Set the animation of shape
shape.Slide.Timeline.MainSequence.AddEffect(shape, wasmModule.AnimationEffectType.Path4PointStar);

//Append new shape - Rectangle and set animation
shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(210, 280, 360, 360)});
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
shape.AppendTextFrame("Animated Shape");
shape.Slide.Timeline.MainSequence.AddEffect(shape, wasmModule.AnimationEffectType.FadedSwivel);

//Append new shape - Cloud and set the animation
shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Cloud,rectangle: wasmModule.RectangleF.FromLTRB(390, 280, 470, 360)});
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_White();
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_CadetBlue();
shape.Slide.Timeline.MainSequence.AddEffect(shape, wasmModule.AnimationEffectType.FadedZoom);
```

---

# Spire.Presentation JavaScript Animation
## Apply animation on chart in PowerPoint presentation
```javascript
//Get the first slide
let slide = ppt.Slides.get_Item(0);
//Get chart
let shape = slide.Shapes.get_Item(0);
if (shape instanceof  wasmModule.IChart){
    //Apply Wipe animation effect to the chart
    let effect = slide.Timeline.MainSequence.AddEffect(shape,  wasmModule.AnimationEffectType.Wipe);
    //Set the BuildType as Series
    effect.GraphicAnimation.BuildType =  wasmModule.GraphicBuildType.BuildAsSeries;
}
```

---

# PowerPoint Shape Animation
## Apply animation effect to a shape in PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Get the first slide
let slide = ppt.Slides.get_Item(0);

//Insert a rectangle in the slide and fill the shape
let shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: wasmModule.RectangleF.FromLTRB(100, 150, 300, 230)});
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_LightBlue();
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
shape.AppendTextFrame("Animated Shape");

//Apply FadedSwivel animation effect to the shape
shape.Slide.Timeline.MainSequence.AddEffect(shape, wasmModule.AnimationEffectType.FadedSwivel);
```

---

# spire.presentation javascript animation
## apply animation on text in PowerPoint
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Add a shape to the slide
let shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: wasmModule.RectangleF.FromLTRB(250, 150, 450, 250)});
shape.Fill.FillType = wasmModule.FillFormatType.Solid;
shape.Fill.SolidColor.Color = wasmModule.Color.get_LightBlue();
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
shape.AppendTextFrame("This demo shows how to apply animation on text in PPT document.");

// Apply animation to the text in shape
let animation = shape.Slide.Timeline.MainSequence.AddEffect(shape, wasmModule.AnimationEffectType.Float);
animation.SetStartEndParagraphs(0, 0);
```

---

# spire.presentation javascript animation
## create custom path animation in PowerPoint
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Add shape
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: wasmModule.RectangleF.FromLTRB(0, 0, 200, 200)});

// Add animation
let effect = ppt.Slides.get_Item(0).Timeline.MainSequence.AddEffect(shape, wasmModule.AnimationEffectType.PathUser);

let common = effect.CommonBehaviorCollection;

let motion = common.get_Item(0);
motion.Origin = wasmModule.AnimationMotionOrigin.Layout;
motion.PathEditMode = wasmModule.AnimationMotionPathEditMode.Relative;

// Add motion path
let moinPath = wasmModule.MotionPath.Create();
moinPath.Add(wasmModule.MotionCommandPathType.MoveTo, wasmModule.PointF.Create(0,0) , wasmModule.MotionPathPointsType.CurveAuto, true);
moinPath.Add(wasmModule.MotionCommandPathType.LineTo, wasmModule.PointF.Create(0.1,0.1), wasmModule.MotionPathPointsType.CurveAuto, true);
moinPath.Add(wasmModule.MotionCommandPathType.LineTo, wasmModule.PointF.Create(-0.1,0.2), wasmModule.MotionPathPointsType.CurveAuto, true);
moinPath.Add(wasmModule.MotionCommandPathType.End, wasmModule.PointF.Create(0,0), wasmModule.MotionPathPointsType.CurveStraight, true);
motion.Path = moinPath;
```

---

# Animation Duration and Delay Time Control
## Get/set duration time or delay time of animations in presentations
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

ppt.LoadFromFile(inputFileName);

//Get the first slide
let slide = ppt.Slides.get_Item(0);
let animations = slide.Timeline.MainSequence;

//Get duration time of animation
let durationTime = animations.get_Item(0).Timing.Duration;

//Set new duration time of animation
animations.get_Item(0).Timing.Duration = 0.8;

//Get delay time of animation
let delayTime = animations.get_Item(0).Timing.TriggerDelayTime;

//Set new delay time of animation
animations.get_Item(0).Timing.TriggerDelayTime = 0.6;
```

---

# Spire.Presentation JavaScript Animation Effect Info
## Extract animation effect information from PowerPoint slides
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

let stringBuilder = [];

// Travel each slide
for(let i = 0; i < ppt.Slides.Count; i++){
    let slide = ppt.Slides.get_Item(i);
    for (let j = 0; j < slide.Timeline.MainSequence.Count; j++){
        let effect = slide.Timeline.MainSequence.get_Item(j);
        // Get the animation effect type
        let animationEffectType = effect.AnimationEffectType;
        stringBuilder.push("animation effect type:" + animationEffectType);

        // Get the slide number where the animation is located
        let slideNumber = slide.SlideNumber;
        stringBuilder.push("slide number:" + slideNumber);

        // Get the shape name
        let shapeName = effect.ShapeTarget.Name;
        stringBuilder.push("shape name:" + shapeName + "\n");
    }
}
```

---

# spire.presentation javascript animation
## get animation motion path from PowerPoint
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

ppt.LoadFromFile(inputFileName);

let slide = ppt.Slides.get_Item(0);

//Get the first shape
let shape = slide.Shapes.get_Item(0);

//Create a StringBuilder to save the tracks
let StringBuilder = [];
let o = 1;
//Traverse all animations
for(let i = 0;i < shape.Slide.Timeline.MainSequence.Count;i++){
    let effect = shape.Slide.Timeline.MainSequence.get_Item(i);
    if (effect.ShapeTarget.Equals(shape)) {
        //Get MotionPath
        let animationMotion = effect.CommonBehaviorCollection.get_Item(0);
        let path = animationMotion.Path;
        for (let j = 0;j < path.Count;j++){
            let motionCmdPath = path.get_Item(j);
            let points = motionCmdPath.Points;
            let type = motionCmdPath.CommandType;
            if(points != null){
                for (let k = 0;k < points.length;k++){
                    let point = points[k];
                    StringBuilder.push(o + "  MotionType: " + type + " -> X: " + point.X + ", Y: " + point.Y);
                }
                o++;
            }
        }
    }
}
```

---

# spire presentation javascript animation
## set animation effect for animate text
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load file
ppt.LoadFromFile(inputFileName);

//Set the AnimateType as Letter
ppt.Slides.get_Item(0).Timeline.MainSequence.get_Item(0).IterateType = wasmModule.AnimateType.Letter;

//Set the IterateTimeValue for the animate text
ppt.Slides.get_Item(0).Timeline.MainSequence.get_Item(0).IterateTimeValue = 10;
```

---

# spire.presentation javascript animation
## set animation repeat type
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load file
ppt.LoadFromFile(inputFileName);

//Get the first slide
let slide = ppt.Slides.get_Item(0);
let animations = slide.Timeline.MainSequence;
animations.get_Item(0).Timing.AnimationRepeatType = wasmModule.AnimationRepeatType.UtilEndOfSlide;
```

---

# Spire.Presentation JavaScript Section Management
## Add sections to PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

//Get the second slide
let slide = ppt.Slides.get_Item(1);

//Append section with section name at the end
ppt.SectionList.Append("E-iceblue01");
//Add section with slide
ppt.SectionList.Add("section1", slide);
```

---

# Spire.Presentation JavaScript Section Operations
## Add slide to section in PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

ppt.LoadFromFile("filename.pptx");

//Add a new shape to the PPT document
ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(200, 50, 500, 150)});

//Create a new section and copy the first slide to it
let NewSection = ppt.SectionList.Append("New Section");
NewSection.Insert(0, ppt.Slides.get_Item(0));
```

---

# Spire.Presentation JavaScript Section Management
## Delete all sections from a PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Remove all sections from the presentation
ppt.SectionList.RemoveAll();
```

---

# Spire.Presentation JavaScript Section Index
## Get the index of a section in a presentation
```javascript
let section = ppt.SectionList.get_Item(0);
//Get the index of the section
let index = ppt.SectionList.IndexOf(section);
```

---

# spire.presentation javascript load from stream
## load encrypted PowerPoint presentation from a stream
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load PowerPoint file from stream
let from_stream =  wasmModule.Stream.CreateByFile(inputFileName);
ppt.LoadFromStream({stream:from_stream, fileFormat:wasmModule.FileFormat.Pptx2013});

// Define the output file name
const outputFileName = "LoadEncryptedStream_out.pptx";

// Save the document to the specified path
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Loop Presentation
## Configure PowerPoint presentation to loop continuously
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load file
ppt.LoadFromFile(inputFileName);

//Set the Boolean value of ShowLoop as true
ppt.ShowLoop = true;

//Set the PowerPoint document to show animation and narration
ppt.ShowAnimation = true;
ppt.ShowNarration = true;
//Use slide transition timings to advance slide
ppt.UseTimings = true;
```

---

# Spire.Presentation JavaScript Page Setup
## Set slide size and orientation in PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Set the size of slides
ppt.SlideSize.Size = wasmModule.SizeF.CreateWH(600, 600);
ppt.SlideSize.Orientation = wasmModule.SlideOrienation.Portrait;
ppt.SlideSize.Type = wasmModule.SlideSizeType.Custom;

let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle, fileName:ImageFileName,rectangle: rect});
ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

// Append new shape
let rec = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 200, 150, (400 + ppt.SlideSize.Size.Width / 2 - 200), 350);
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle: rec});
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
shape.Fill.FillType = wasmModule.FillFormatType.None;

// Add text to shape
shape.AppendTextFrame("The sample demonstrates how to set slide size.");

shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Myriad Pro");
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).FontHeight = 24;
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.FromArgb(36, 64, 97);
```

---

# Save PowerPoint Presentation to Stream
## Demonstrates how to save a PowerPoint document to a stream using Spire.Presentation for JavaScript
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Define the output file name
const outputFileName = "SaveToStream_out.pptx";

//Save to Stream
let to_stream = wasmModule.Stream.CreateByFile(outputFileName);

ppt.SaveToFile({stream:to_stream,fileFormat:wasmModule.FileFormat.Pptx2013});

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript ShowType
## Set presentation show type as kiosk
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load file
ppt.LoadFromFile("InputTemplate.pptx");

// Specify the presentation show type as kiosk
ppt.ShowType = wasmModule.SlideShowType.Kiosk;

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Document Operation
## Split PowerPoint presentation into individual slides
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load file
ppt.LoadFromFile(inputFileName);

for (let i = 0; i < ppt.Slides.Count; i++) {
  // Initialize another instance of Presentation, and remove the blank slide
  let newppt = wasmModule.Presentation.Create();
  newppt.Slides.RemoveAt(0);

  // Append the specified slide from old presentation to the new one
  newppt.Slides.Append({slide:ppt.Slides.get_Item(i)});

  // Save the document
  let outputFileName = `SplitPPT-${i}.pptx`;
  newppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });
  newppt.Dispose();
}

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Properties
## Get Built-in Properties from PowerPoint Document
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load the PPT document from disk
ppt.LoadFromFile(inputFileName);

// Get the builtin properties
let application = ppt.DocumentProperty.Application;
let author = ppt.DocumentProperty.Author;
let company = ppt.DocumentProperty.Company;
let keywords = ppt.DocumentProperty.Keywords;
let comments = ppt.DocumentProperty.Comments;
let category = ppt.DocumentProperty.Category;
let title = ppt.DocumentProperty.Title;
let subject = ppt.DocumentProperty.Subject;

// Create content to save
let content = [];
content.push("DocumentProperty.Application: " + application);
content.push("DocumentProperty.Author: " + author);
content.push("DocumentProperty.Company " + company);
content.push("DocumentProperty.Keywords: " + keywords);
content.push("DocumentProperty.Comments: " + comments);
content.push("DocumentProperty.Category: " + category);
content.push("DocumentProperty.Title: " + title);
content.push("DocumentProperty.Subject: " + subject);
```

---

# Spire.Presentation JavaScript MarkAsFinal
## Mark PowerPoint document as final
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the document from disk
ppt.LoadFromFile(inputFileName);

//Mark the document as final
ppt.DocumentProperty["_MarkAsFinal"] = true;
```

---

# spire presentation document properties
## set document properties of a PPT file
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

ppt.LoadFromFile(inputFileName);

// Set the DocumentProperty of PPT document
ppt.DocumentProperty.Application = "Spire.Presentation";
ppt.DocumentProperty.Author = "E-iceblue";
ppt.DocumentProperty.Company = "E-iceblue Co., Ltd.";
ppt.DocumentProperty.Keywords = "Demo File";
ppt.DocumentProperty.Comments = "This file is used to test Spire.Presentation.";
ppt.DocumentProperty.Category = "Demo";
ppt.DocumentProperty.Title = "This is a demo file.";
ppt.DocumentProperty.Subject = "Test";

// Save the document to the specified path
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Template Properties
## Set document properties for a presentation template
```javascript
// Create a document
let ppt = wasmModule.Presentation.Create();

// Set the DocumentProperty
ppt.DocumentProperty.Application = "Spire.Presentation";
ppt.DocumentProperty.Author = "E-iceblue";
ppt.DocumentProperty.Company = "E-iceblue Co., Ltd.";
ppt.DocumentProperty.Keywords = "Demo File";
ppt.DocumentProperty.Comments = "This file is used to test Spire.Presentation.";
ppt.DocumentProperty.Category = "Demo";
ppt.DocumentProperty.Title = "This is a demo file.";
ppt.DocumentProperty.Subject = "Test";

// Save to template file
ppt.SaveToFile({file:fileName,fileFormat:fileFormat});
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Security
## Check if a PPT document is password protected
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Check whether a PPT document is password protected
let isProtected = ppt.IsPasswordProtected(inputFileName);
let outString = "The file is " + (isProtected ? "password " : "not password ") + "protected!";
```

---

# PowerPoint Document Encryption
## Encrypt a PPT document with a password
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Get the password that the user entered
let password = "e-iceblue";

//Encrypy the document with the password
ppt.Encrypt(password);

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Security
## Modify password of encrypted PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load the file from disk with password
ppt.LoadFromFile({file:inputFileName, password:"123456"});

// Remove the encryption
ppt.RemoveEncryption();

// Protect the document by setting a new password
ppt.Protect("654321");

// Save the document to the specified path
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# Open Encrypted PowerPoint Presentation
## Demonstrates how to open an encrypted PPT document with password
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the PPT with password
ppt.LoadFromFile({file:inputFileName, fileFormat:wasmModule.FileFormat.Pptx2010,password: "123456"});
```

---

# spire.presentation javascript security
## remove all digital signatures from powerpoint
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load the file from disk
ppt.LoadFromFile(inputFileName);

// Remove all digital signatures
if (ppt.IsDigitallySigned == true) {
    ppt.RemoveAllDigitalSignatures();
}

// Save the document to the specified path
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Encryption Removal
## Remove password protection from PowerPoint presentations
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the password-protected file
ppt.LoadFromFile({file: "Template_Ppt_4.pptx", password: "123456"});

//Remove encryption
ppt.RemoveEncryption();

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Security
## Set PowerPoint document to read-only by protecting with password
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the document from disk
ppt.LoadFromFile(inputFileName);

//Get the password that the user entered
let password = "e-iceblue";

//Protect the document with the password
ppt.Protect(password);

// Define the output file name
const outputFileName = "SetDocumentReadOnly_out.pptx";

// Save the document to the specified path
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Background
## Set different background styles for presentation slides
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

ppt.LoadFromFile(inputFileName);

//Set the background of the first slide to Gradient color
ppt.Slides.get_Item(0).SlideBackground.Type = wasmModule.BackgroundType.Custom;
ppt.Slides.get_Item(0).SlideBackground.Fill.FillType = wasmModule.FillFormatType.Gradient;
ppt.Slides.get_Item(0).SlideBackground.Fill.Gradient.GradientShape = wasmModule.GradientShapeType.Linear;
ppt.Slides.get_Item(0).SlideBackground.Fill.Gradient.GradientStyle = wasmModule.GradientStyle.FromCorner1;
ppt.Slides.get_Item(0).SlideBackground.Fill.Gradient.GradientStops.Append({position:1,knownColor: wasmModule.KnownColors.SkyBlue});
ppt.Slides.get_Item(0).SlideBackground.Fill.Gradient.GradientStops.Append({position:0,knownColor: wasmModule.KnownColors.White});

//Set the background of the second slide to Solid color
ppt.Slides.get_Item(1).SlideBackground.Type = wasmModule.BackgroundType.Custom;
ppt.Slides.get_Item(1).SlideBackground.Fill.FillType = wasmModule.FillFormatType.Solid;
ppt.Slides.get_Item(1).SlideBackground.Fill.SolidColor.Color = wasmModule.Color.get_SkyBlue();

ppt.Slides.Append();

//Set the background of the third slide to picture
let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
ppt.Slides.get_Item(2).SlideBackground.Fill.FillType = wasmModule.FillFormatType.Picture;
let image = ppt.Slides.get_Item(2).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName: ImageFileName,rectangle: rect});
ppt.Slides.get_Item(2).SlideBackground.Fill.PictureFill.Picture.EmbedImage = image.PictureFill.Picture.EmbedImage;
```

---

# Spire Presentation JavaScript Gradient Background
## Set gradient background for slides
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load document from disk
ppt.LoadFromFile(inputFileName);

// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Set the background to gradient
slide.SlideBackground.Type = wasmModule.BackgroundType.Custom;
slide.SlideBackground.Fill.FillType = wasmModule.FillFormatType.Gradient;

// Add gradient stops
slide.SlideBackground.Fill.Gradient.GradientStops.Append({position:0.1,color: wasmModule.Color.get_LightSeaGreen()});
slide.SlideBackground.Fill.Gradient.GradientStops.Append({position:0.7,color: wasmModule.Color.get_LightCyan()});

// Set gradient shape type
slide.SlideBackground.Fill.Gradient.GradientShape = wasmModule.GradientShapeType.Linear;

// Set the angle
slide.SlideBackground.Fill.Gradient.LinearGradientFill.Angle = 45;
```

---

# spire presentation master background
## set solid background for master slide
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Set the slide background of master
ppt.Masters.get_Item(0).SlideBackground.Type = wasmModule.BackgroundType.Custom;
ppt.Masters.get_Item(0).SlideBackground.Fill.FillType = wasmModule.FillFormatType.Solid;
ppt.Masters.get_Item(0).SlideBackground.Fill.SolidColor.Color = wasmModule.Color.get_LightSalmon();
```

---

# Spire.Presentation JavaScript Error Bars
## Add and format error bars in PowerPoint charts
```javascript
//Get the column chart on the first slide and set chart title.
let columnChart = ppt.Slides.get_Item(0).Shapes.get_Item(0);
columnChart.ChartTitle.TextProperties.Text = "Vertical Error Bars";

//Add Y (Vertical) Error Bars.
//Get Y error bars of the first chart series.
let errorBarsYFormat1 = columnChart.Series.get_Item(0).ErrorBarsYFormat;

//Set end cap.
errorBarsYFormat1.ErrorBarNoEndCap = false;

//Specify direction.
errorBarsYFormat1.ErrorBarSimType = wasmModule.ErrorBarSimpleType.Plus;

//Specify error amount type.
errorBarsYFormat1.ErrorBarvType = wasmModule.ErrorValueType.StandardError;

//Set value.
errorBarsYFormat1.ErrorBarVal = 0.3;

//Set line format.
errorBarsYFormat1.Line.FillType = wasmModule.FillFormatType.Solid;
errorBarsYFormat1.Line.SolidFillColor.Color = wasmModule.Color.get_MediumVioletRed();
errorBarsYFormat1.Line.Width = 1;

//Get the bubble chart on the second slide and set chart title.
let bubbleChart = ppt.Slides.get_Item(1).Shapes.get_Item(0);
bubbleChart.ChartTitle.TextProperties.Text = "Vertical and Horizontal Error Bars";

//Add X (Horizontal) and Y (Vertical) Error Bars.
//Get X error bars of the first chart series.
let errorBarsXFormat = bubbleChart.Series.get_Item(0).ErrorBarsXFormat;

//Set end cap.
errorBarsXFormat.ErrorBarNoEndCap = false;

//Specify direction.
errorBarsXFormat.ErrorBarSimType = wasmModule.ErrorBarSimpleType.Both;

//Specify error amount type.
errorBarsXFormat.ErrorBarvType = wasmModule.ErrorValueType.StandardError;

//Set value.
errorBarsXFormat.ErrorBarVal = 0.3;

//Get Y error bars of the first chart series.
let errorBarsYFormat2 = bubbleChart.Series.get_Item(0).ErrorBarsYFormat;

//Set end cap.
errorBarsYFormat2.ErrorBarNoEndCap = false;

//Specify direction.
errorBarsYFormat2.ErrorBarSimType = wasmModule.ErrorBarSimpleType.Both;

//Specify error amount type.
errorBarsYFormat2.ErrorBarvType = wasmModule.ErrorValueType.StandardError;

//Set value.
errorBarsYFormat2.ErrorBarVal = 0.3;
```

---

# spire.presentation javascript chart
## add custom error bars to chart
```javascript
//Get the bubble chart on the first slide
let bubbleChart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Get X error bars of the first chart series
let errorBarsXFormat = bubbleChart.Series.get_Item(0).ErrorBarsXFormat;
//Specify error amount type as custom error bars
errorBarsXFormat.ErrorBarvType = wasmModule.ErrorValueType.CustomErrorBars;
//Set the minus and plus value of the X error bars
errorBarsXFormat.MinusVal = 0.5;
errorBarsXFormat.PlusVal = 0.5;

//Get Y error bars of the first chart series
let errorBarsYFormat = bubbleChart.Series.get_Item(0).ErrorBarsYFormat;
//Specify error amount type as custom error bars
errorBarsYFormat.ErrorBarvType = wasmModule.ErrorValueType.CustomErrorBars;
//Set the minus and plus value of the Y error bars
errorBarsYFormat.MinusVal = 1;
errorBarsYFormat.PlusVal = 1;
```

---

# spire presentation javascript chart
## add secondary value axis to chart
```javascript
//Get the chart from the PowerPoint file.
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Add a secondary axis to display the value of Series 3.
chart.Series.get_Item(2).UseSecondAxis = true;

//Set the grid line of secondary axis as invisible.
chart.SecondaryValueAxis.MajorGridTextLines.FillType = wasmModule.FillFormatType.None;
```

---

# spire presentation javascript chart
## add shadow effect to chart data labels
```javascript
//Get the chart.
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Add a data label to the first chart series.
let dataLabels = chart.Series.get_Item(0).DataLabels;
let Label = dataLabels.Add();
Label.LabelValueVisible = true;

//Add outer shadow effect to the data label.
Label.Effect.OuterShadowEffect = wasmModule.OuterShadowEffect.Create();

//Set shadow color.
Label.Effect.OuterShadowEffect.ColorFormat.Color = wasmModule.Color.get_Yellow();

//Set blur.
Label.Effect.OuterShadowEffect.BlurRadius = 5;

//Set distance.
Label.Effect.OuterShadowEffect.Distance = 10;

//Set angle.
Label.Effect.OuterShadowEffect.Direction = 90;
```

---

# spire presentation javascript trendline
## add trend line for chart series
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Get the target chart, add trendline for the first data series of the chart and specify the trendline type.
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);
let it = chart.Series.get_Item(0).AddTrendLine(wasmModule.TrendlinesType.Linear);

//Set the trendline properties to determine what should be displayed.
it.displayEquation = false;
it.displayRSquaredValue = false;
```

---

# spire.presentation javascript chart
## Auto vary color for pie chart
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

let rect1 = wasmModule.RectangleF.FromLTRB(40, 100, 590, 420);
// Add a pie chart
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Pie, rectangle: rect1, init: false });
chart.ChartTitle.TextProperties.Text = "Sales by Quarter";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;
// Attach the data to chart
let quarters = ["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"];
let sales = [210, 320, 180, 500];
chart.ChartData._get_Item(0, 0).Text = "Quarters";
chart.ChartData._get_Item(0, 1).Text = "Sales";
for (let i = 0; i < quarters.length; ++i) {
  chart.ChartData._get_Item(i + 1, 0).Text = quarters[i];
  chart.ChartData._get_Item(i + 1, 1).NumberValue = sales[i];
}

chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("B1", "B1");
chart.Categories.CategoryLabels = chart.ChartData._get_ItemNE("A2", "A5");
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemNE("B2", "B5");

// Set whether auto vary color, default value is true
chart.Series.get_Item(0).IsVaryColor = false;
chart.Series.get_Item(0).Distance = 15;
```

---

# spire.presentation javascript chart legend
## change color for chart legend
```javascript
//Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Change the fill color
Chart.ChartLegend.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.Fill.FillType = wasmModule.FillFormatType.Solid;
Chart.ChartLegend.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.Fill.SolidColor.Color = wasmModule.Color.get_Blue();
//Use italic for the paragraph
Chart.ChartLegend.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.IsItalic = wasmModule.TriState.True;
```

---

# spire presentation javascript chart font
## change font size for chart data table
```javascript
//Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);
Chart.HasDataTable = true;

//Add a new paragraph in data table
Chart.ChartDataTable.Text.Paragraphs._Append(wasmModule.TextParagraph.Create());
//Change the font size
Chart.ChartDataTable.Text.Paragraphs.get_Item(0).DefaultCharacterProperties.FontHeight = 15;
```

---

# Spire.Presentation JavaScript Chart Legend
## Change font size for chart legend
```javascript
//Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Change legend font size
Chart.ChartLegend.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.FontHeight = 17;
```

---

# Spire.Presentation JavaScript Chart Series Name
## Change chart series name in PowerPoint presentation
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

// Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Get the ranges of series label
let cr = Chart.Series.SeriesLabel;

// Change the value
cr.get_Item(0).Text = "Changed series name";
```

---

# Spire.Presentation JavaScript TrendLine
## Change font size and position for TrendLine equation
```javascript
//Get chart on the first slide
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Get the first trendline
let trendline = chart.Series.get_Item(0).TrendLines[0];

//Change font size for trendline Equation text
for (let i = 0; i < trendline.TrendLineLabel.TextFrameProperties.Paragraphs.Count; i++) {
  let para = trendline.TrendLineLabel.TextFrameProperties.Paragraphs.get_Item(i);
  para.DefaultCharacterProperties.FontHeight = 20;
  for (let j = 0; j < para.TextRanges.Count; j++) {
    let range = para.TextRanges.get_Item(j);
    range.FontHeight = 20;
  }
}

//Change position for trendline Equation
trendline.TrendLineLabel.OffsetX = -0.1;
trendline.TrendLineLabel.OffsetY = -0.05;
```

---

# spire.presentation javascript chart
## change text font in chart
```javascript
//Get the chart
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Change the font of title
chart.ChartTitle.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");
chart.ChartTitle.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Blue;
chart.ChartTitle.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.FontHeight = 30;

//Change the font of legend
chart.ChartLegend.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.Fill.SolidColor.KnownColor = wasmModule.KnownColors.DarkGreen;
chart.ChartLegend.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");

//Change the font of series
chart.PrimaryCategoryAxis.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Red;
chart.PrimaryCategoryAxis.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.Fill.FillType = wasmModule.FillFormatType.Solid;
chart.PrimaryCategoryAxis.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.FontHeight = 10;
chart.PrimaryCategoryAxis.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");
```

---

# Chart Axis Configuration in Spire.Presentation
## Configuring primary and secondary axes in a PowerPoint chart
```javascript
//Get the chart
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Add a secondary axis to display the value of Series 3
chart.Series.get_Item(2).UseSecondAxis = true;

//Set the grid line of secondary axis as invisible
chart.SecondaryValueAxis.MajorGridTextLines.FillType = wasmModule.FillFormatType.None;

//Set bounds of axis value. Before we assign values, we must set IsAutoMax and IsAutoMin as false, otherwise MS PowerPoint will automatically set the values.
chart.PrimaryValueAxis.IsAutoMax = false;
chart.PrimaryValueAxis.IsAutoMin = false;
chart.SecondaryValueAxis.IsAutoMax = false;
chart.SecondaryValueAxis.IsAutoMax = false;

chart.PrimaryValueAxis.MinValue = 0;
chart.PrimaryValueAxis.MaxValue = 5.0;
chart.SecondaryValueAxis.MinValue = 0;
chart.SecondaryValueAxis.MaxValue = 1.0;

//Set axis line format
chart.PrimaryValueAxis.MinorGridLines.FillType = wasmModule.FillFormatType.Solid;
chart.SecondaryValueAxis.MinorGridLines.FillType = wasmModule.FillFormatType.Solid;
chart.PrimaryValueAxis.MinorGridLines.Width = 0.1;
chart.SecondaryValueAxis.MinorGridLines.Width = 0.1;
chart.PrimaryValueAxis.MinorGridLines.SolidFillColor.Color = wasmModule.Color.get_LightGray();
chart.SecondaryValueAxis.MinorGridLines.SolidFillColor.Color = wasmModule.Color.get_LightGray();
chart.PrimaryValueAxis.MinorGridLines.DashStyle = wasmModule.LineDashStyleType.Dash;
chart.SecondaryValueAxis.MinorGridLines.DashStyle = wasmModule.LineDashStyleType.Dash;

chart.PrimaryValueAxis.MajorGridTextLines.Width = 0.3;
chart.PrimaryValueAxis.MajorGridTextLines.SolidFillColor.Color = wasmModule.Color.get_LightSkyBlue();
chart.SecondaryValueAxis.MajorGridTextLines.Width = 0.3;
chart.SecondaryValueAxis.MajorGridTextLines.SolidFillColor.Color = wasmModule.Color.get_LightSkyBlue();
```

---

# Spire.Presentation JavaScript Chart Copying
## Copy a chart between PowerPoint presentations
```javascript
// Get the chart that is going to be copied
let chart = ppt1.Slides.get_Item(0).Shapes.get_Item(0);

// Copy chart from the first document to the second document
ppt2.Slides.Append();
ppt2.Slides.get_Item(1).Shapes.CreateChart(chart, wasmModule.RectangleF.FromLTRB(100, 100, 600, 400), -1);
```

---

# Spire.Presentation JavaScript Chart Copy
## Copy a chart from one slide to another within the same PowerPoint presentation
```javascript
//Get the chart that is going to be copied.
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Copy the chart from the first slide to the specified location of the second slide within the same document.
let slide1 = ppt.Slides.Append();
slide1.Shapes.CreateChart(chart, wasmModule.RectangleF.FromLTRB(100, 100, 600, 400), 0);
```

---

# Spire.Presentation JavaScript Chart
## Create 100% Stacked Bar Chart
```javascript
//Create a PowerPoint document.
let ppt = wasmModule.Presentation.Create();

//Add a "Bar100PercentStacked" chart to the first slide.
ppt.SlideSize.Type = wasmModule.SlideSizeType.Screen16x9;
let slidesize = ppt.SlideSize.Size;

let slide = ppt.Slides.get_Item(0);

//Append a chart.
let rect = wasmModule.RectangleF.FromLTRB(20, 20, slidesize.Width - 20, slidesize.Height - 20);
let chart = slide.Shapes.AppendChart({ type: wasmModule.ChartType.Bar100PercentStacked, rectangle: rect });

//Write data to the chart data.
let columnlabels = ["Series 1", "Series 2", "Series 3"];

//Insert the column labels.
for (let i = 0; i < columnlabels.length; i++) {
  chart.ChartData._get_Item(0, i + 1).Text = columnlabels[i];
}

let rowlabels = ["Category 1", "Category 2", "Category 3"];

//Insert the row labels.
for (let i = 0; i < rowlabels.length; i++) {
  chart.ChartData._get_Item(i + 1, 0).Text = rowlabels[i];
}

let values = [[20.83233, 10.34323, -10.354667], [10.23456, -12.23456, 23.34456], [12.34345, -23.34343, -13.23232]];

//Insert the values.
let value = 0.0;
for (let i = 0; i < rowlabels.length; i++) {
  for (let j = 0; j < columnlabels.length; j++) {
    value = Math.round(values[i][j], 2);
    chart.ChartData._get_Item(i + 1, j + 1).NumberValue = value;
  }
}

chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 1, 0, columnlabels.length);
chart.Categories.CategoryLabels = chart.ChartData._get_ItemRCLL(1, 0, rowlabels.length, 0);

//Set the position of category axis.
chart.PrimaryCategoryAxis.Position = wasmModule.AxisPositionType.Left;
chart.SecondaryCategoryAxis.Position = wasmModule.AxisPositionType.Left;
chart.PrimaryCategoryAxis.TickLabelPosition = wasmModule.TickLabelPositionType.TickLabelPositionLow;

//Set the data, font and format for the series of each column.
for (let i = 0; i < columnlabels.length; i++) {
  chart.Series.get_Item(i).Values = chart.ChartData._get_ItemRCLL(1, i + 1, rowlabels.length, i + 1);
  chart.Series.get_Item(i).Fill.FillType = wasmModule.FillFormatType.Solid;
  chart.Series.get_Item(i).InvertIfNegative = false;
  for (let j = 0; j < rowlabels.length; j++) {
    let label = chart.Series.get_Item(i).DataLabels.Add();
    label.LabelValueVisible = true;
    chart.Series.get_Item(i).DataLabels.get_Item(j).HasDataSource = false;
    chart.Series.get_Item(i).DataLabels.get_Item(j).NumberFormat = "0#\\%";
    chart.Series.get_Item(i).DataLabels.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.FontHeight = 12;
  }
}

//Set the color of the Series.
chart.Series.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_YellowGreen();
chart.Series.get_Item(1).Fill.SolidColor.Color = wasmModule.Color.get_Red();
chart.Series.get_Item(2).Fill.SolidColor.Color = wasmModule.Color.get_Green();

let font = wasmModule.TextFont.Create("Tw Cen MT");

//Set the font and size for chartlegend.
for (let k = 0; k < chart.ChartLegend.EntryTextProperties.length; k++) {
  let textPara = chart.ChartLegend.EntryTextProperties[k];
  textPara.LatinFont = font;
  textPara.FontHeight = 20;
}
```

---

# spire presentation javascript box and whisker chart
## create a Box and Whisker chart in PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Insert a BoxAndWhisker chart to the first slide
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.BoxAndWhisker, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 550, 450), init: false });

// Series labels
let seriesLabel = ["Series 1", "Series 2", "Series 3"];
for (let i = 0; i < seriesLabel.length; i++) {
  chart.ChartData._get_Item(0, i + 1).Text = "Series 1";
}

// Categories
let categories = ["Category 1", "Category 1", "Category 1", "Category 1", "Category 1", "Category 1", "Category 1",
  "Category 2", "Category 2", "Category 2", "Category 2", "Category 2", "Category 2",
  "Category 3", "Category 3", "Category 3", "Category 3", "Category 3"];
for (let i = 0; i < categories.length; i++) {
  chart.ChartData._get_Item(i + 1, 0).Text = categories[i];
}

// Values
let values = [[-7, -3, -24], [-10, 1, 11], [-28, -6, 34], [47, 2, -21], [35, 17, 22], [-22, 15, 19], [17, -11, 25],
[-30, 18, 25], [49, 22, 56], [37, 22, 15], [-55, 25, 31], [14, 18, 22], [18, -22, 36], [-45, 25, -17],
[-33, 18, 22], [18, 2, -23], [-33, -22, 10], [10, 19, 22]];
for (let i = 0; i < seriesLabel.length; i++) {
  for (let j = 0; j < categories.length; j++) {
    chart.ChartData._get_Item(j + 1, i + 1).NumberValue = values[j][i];
  }
}

//Set series
chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 1, 0, seriesLabel.length);
chart.Categories.CategoryLabels = chart.ChartData._get_ItemRCLL(1, 0, categories.length, 0);
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemRCLL(1, 1, categories.length, 1);
chart.Series.get_Item(1).Values = chart.ChartData._get_ItemRCLL(1, 2, categories.length, 2);
chart.Series.get_Item(2).Values = chart.ChartData._get_ItemRCLL(1, 3, categories.length, 3);
chart.Series.get_Item(0).ShowInnerPoints = false;
chart.Series.get_Item(0).ShowOutlierPoints = true;
chart.Series.get_Item(0).ShowMeanMarkers = true;
chart.Series.get_Item(0).ShowMeanLine = true;
chart.Series.get_Item(0).QuartileCalculationType = wasmModule.QuartileCalculation.ExclusiveMedian;
chart.Series.get_Item(1).ShowInnerPoints = false;
chart.Series.get_Item(1).ShowOutlierPoints = true;
chart.Series.get_Item(1).ShowMeanMarkers = true;
chart.Series.get_Item(1).ShowMeanLine = true;
chart.Series.get_Item(1).QuartileCalculationType = wasmModule.QuartileCalculation.InclusiveMedian;
chart.Series.get_Item(2).ShowInnerPoints = false;
chart.Series.get_Item(2).ShowOutlierPoints = true;
chart.Series.get_Item(2).ShowMeanMarkers = true;
chart.Series.get_Item(2).ShowMeanLine = true;
chart.Series.get_Item(2).QuartileCalculationType = wasmModule.QuartileCalculation.ExclusiveMedian;

//Show legend
chart.HasLegend = true;
chart.ChartTitle.TextProperties.Text = "BoxAndWhisker";
chart.ChartLegend.Position = wasmModule.ChartLegendPositionType.Top;
```

---

# Spire.Presentation JavaScript Bubble Chart
## Create a bubble chart in a PowerPoint presentation
```javascript
// Create a PPT file
let ppt = wasmModule.Presentation.Create();

// Add bubble chart
let rect1 = wasmModule.RectangleF.FromLTRB(90, 100, 640, 420);
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Bubble, rectangle: rect1, init: false });

// Chart title
chart.ChartTitle.TextProperties.Text = "Bubble Chart";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

// Set chart data headers
chart.ChartData._get_Item(0, 0).Text = "X-Value";
chart.ChartData._get_Item(0, 1).Text = "Y-Value";
chart.ChartData._get_Item(0, 2).Text = "Size";

// Set series label
chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("B1", "B1");

chart.Series.get_Item(0).XValues = chart.ChartData._get_ItemNE("A2", "A5");
chart.Series.get_Item(0).YValues = chart.ChartData._get_ItemNE("B2", "B5");
chart.Series.get_Item(0).Bubbles.Add({ cellRange: chart.ChartData._get_ItemN("C2") });
chart.Series.get_Item(0).Bubbles.Add({ cellRange: chart.ChartData._get_ItemN("C3") });
chart.Series.get_Item(0).Bubbles.Add({ cellRange: chart.ChartData._get_ItemN("C4") });
chart.Series.get_Item(0).Bubbles.Add({ cellRange: chart.ChartData._get_ItemN("C5") });
```

---

# spire.presentation javascript chart
## create clustered column chart in PowerPoint presentation
```javascript
//Create a PPT file
let ppt = wasmModule.Presentation.Create();

//Add clustered column chart
let rect1 = wasmModule.RectangleF.FromLTRB(90, 100, 640, 420);
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.ColumnClustered, rectangle: rect1, init: false });

//Chart title
chart.ChartTitle.TextProperties.Text = "Clustered Column Chart";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Set series text
chart.ChartData._get_Item(0, 1).Text = "Series1";
chart.ChartData._get_Item(0, 2).Text = "Series2";

//Set category text
chart.ChartData._get_Item(1, 0).Text = "Category 1";
chart.ChartData._get_Item(2, 0).Text = "Category 2";
chart.ChartData._get_Item(3, 0).Text = "Category 3";
chart.ChartData._get_Item(4, 0).Text = "Category 4";

//Set series label
chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("B1", "C1");
//Set category label
chart.Categories.CategoryLabels = chart.ChartData._get_ItemNE("A2", "A5");

//Set values for series
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemNE("B2", "B5");
chart.Series.get_Item(1).Values = chart.ChartData._get_ItemNE("C2", "C5");
```

---

# Spire.Presentation JavaScript Combination Chart
## Create a combination chart with column and line series in a PowerPoint presentation
```javascript
//Create a presentation instance
let ppt = wasmModule.Presentation.Create();

//Insert a column clustered chart
let rect = wasmModule.RectangleF.FromLTRB(100, 100, 650, 420);
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.ColumnClustered, rectangle: rect });

//Set chart title
chart.ChartTitle.TextProperties.Text = "Monthly Sales Report";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Create a datatable
let caption = ["Month", "Sales", "Growth rate"];
let month = ["January", "February", "March", "April", "May", "June"];
let sales = [200, 250, 300, 150, 200, 400];
let growth_rate = [0.6, 0.8, 0.6, 0.2, 0.5, 0.9];

//Import data from datatable to chart data
for (let i = 0; i < caption.length; i++) {
  chart.ChartData._get_Item(0, i).Text = caption[i];
}
for (let i = 0; i < month.length; i++) {
  chart.ChartData._get_Item(i + 1, 0).Text = month[i];
}
for (let i = 0; i < sales.length; i++) {
  chart.ChartData._get_Item(i + 1, 1).NumberValue = sales[i];
}
for (let i = 0; i < growth_rate.length; i++) {
  chart.ChartData._get_Item(i + 1, 2).NumberValue = growth_rate[i];
}

//Set series labels
chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("B1", "C1");

//Set categories labels
chart.Categories.CategoryLabels = chart.ChartData._get_ItemNE("A2", "A7");

//Assign data to series values
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemNE("B2", "B7");
chart.Series.get_Item(1).Values = chart.ChartData._get_ItemNE("C2", "C7");

//Change the chart type of serie 2 to line with markers
chart.Series.get_Item(1).Type = wasmModule.ChartType.LineMarkers;

//Plot data of series 2 on the secondary axis
chart.Series.get_Item(1).UseSecondAxis = true;

//Set the number format as percentage
chart.SecondaryValueAxis.NumberFormat = "0%";

//Hide gridlinkes of secondary axis
chart.SecondaryValueAxis.MajorGridTextLines.FillType = wasmModule.FillFormatType.None;

//Set overlap
chart.OverLap = -50;

//Set gapwidth
chart.GapWidth = 200;
```

---

# spire.presentation javascript chart
## create 3D cylinder clustered chart in PowerPoint
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Set background image
let rect2 = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: inputFileName, rectangle: rect2 });
ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

//Insert chart
let rect = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 200, 85, (400 + ppt.SlideSize.Size.Width / 2 - 200), 485);
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Cylinder3DClustered, rectangle: rect });

//Add chart Title
chart.ChartTitle.TextProperties.Text = "Report";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Load data from datatable to chart
chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("B1", "D1");
chart.Categories.CategoryLabels = chart.ChartData._get_ItemNE("A2", "A7");
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemNE("B2", "B7");
chart.Series.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
chart.Series.get_Item(0).Fill.SolidColor.KnownColor = wasmModule.KnownColors.Brown;
chart.Series.get_Item(1).Values = chart.ChartData._get_ItemNE("C2", "C7");
chart.Series.get_Item(1).Fill.FillType = wasmModule.FillFormatType.Solid;
chart.Series.get_Item(1).Fill.SolidColor.KnownColor = wasmModule.KnownColors.Green;
chart.Series.get_Item(2).Values = chart.ChartData._get_ItemNE("D2", "D7");
chart.Series.get_Item(2).Fill.FillType = wasmModule.FillFormatType.Solid;
chart.Series.get_Item(2).Fill.SolidColor.KnownColor = wasmModule.KnownColors.Orange;

//Set the 3D rotation
chart.RotationThreeD.XDegree = 10;
chart.RotationThreeD.YDegree = 10;
```

---

# spire.presentation javascript chart
## create doughnut chart in PowerPoint presentation
```javascript
// Create a ppt document
let ppt = wasmModule.Presentation.Create();

let rect = wasmModule.RectangleF.FromLTRB(80, 100, 630, 420);

// Set background image
let rect2 = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: inputFileName, rectangle: rect2 });
ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

// Add a Doughnut chart
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Doughnut, rectangle: rect, init: false });
chart.ChartTitle.TextProperties.Text = "Market share by country";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;

let countries = ["Guba", "Mexico", "France", "German"];
let sales = [1800, 3000, 5100, 6200];
chart.ChartData._get_Item(0, 0).Text = "Countries";
chart.ChartData._get_Item(0, 1).Text = "Sales";
for (let i = 0; i < countries.length; ++i) {
  chart.ChartData._get_Item(i + 1, 0).Text = countries[i];
  chart.ChartData._get_Item(i + 1, 1).NumberValue = sales[i];
}
chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("B1", "B1");
chart.Categories.CategoryLabels = chart.ChartData._get_ItemNE("A2", "A5");
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemNE("B2", "B5");

for (let i = 0; i < chart.Series.get_Item(0).Values.Count; i++) {
  let cdp = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
  cdp.Index = i;
  chart.Series.get_Item(0).DataPoints.Add(cdp);
}
// Set the series color
chart.Series.get_Item(0).DataPoints.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
chart.Series.get_Item(0).DataPoints.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_LightBlue();
chart.Series.get_Item(0).DataPoints.get_Item(1).Fill.FillType = wasmModule.FillFormatType.Solid;
chart.Series.get_Item(0).DataPoints.get_Item(1).Fill.SolidColor.Color = wasmModule.Color.get_MediumPurple();
chart.Series.get_Item(0).DataPoints.get_Item(2).Fill.FillType = wasmModule.FillFormatType.Solid;
chart.Series.get_Item(0).DataPoints.get_Item(2).Fill.SolidColor.Color = wasmModule.Color.get_DarkGray();
chart.Series.get_Item(0).DataPoints.get_Item(3).Fill.FillType = wasmModule.FillFormatType.Solid;
chart.Series.get_Item(0).DataPoints.get_Item(3).Fill.SolidColor.Color = wasmModule.Color.get_DarkOrange();

chart.Series.get_Item(0).DataLabels.LabelValueVisible = true;
chart.Series.get_Item(0).DataLabels.PercentValueVisible = true;
chart.Series.get_Item(0).DoughnutHoleSize = 60;
```

---

# Spire.Presentation JavaScript Funnel Chart
## Create a Funnel Chart in a PowerPoint Presentation
```javascript
//Create PPT document
let ppt = wasmModule.Presentation.Create();

//Create a Funnel chart to the first slide
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Funnel, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 600, 450), init: false });

//Set series text
chart.ChartData._get_Item(0, 1).Text = "Series 1";

//Set category text
let categories = ["Website Visits", "Download", "Uploads", "Requested price", "Invoice sent", "Finalized"];
for (let i = 0; i < categories.length; i++) {
  chart.ChartData._get_Item(i + 1, 0).Text = categories[i];
}

//Fill data for chart
let values = [50000, 47000, 30000, 15000, 9000, 5600];
for (let i = 0; i < values.length; i++) {
  chart.ChartData._get_Item(i + 1, 1).NumberValue = values[i];
}

//Set series labels
chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 1, 0, 1);

//Set categories labels
chart.Categories.CategoryLabels = chart.ChartData._get_ItemRCLL(1, 0, categories.length, 0);

//Assign data to series values
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemRCLL(1, 1, values.length, 1);

//Set the chart title
chart.ChartTitle.TextProperties.Text = "Funnel";
```

---

# Spire.Presentation JavaScript Chart
## Create Histogram Chart
```javascript
//Create PPT document
let ppt = wasmModule.Presentation.Create();

//Add a Histogram chart
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Histogram, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 550, 450), init: false });

//Set series text
chart.ChartData._get_Item(0, 0).Text = "Series 1";

//Fill data for chart
let values = [1, 1, 1, 3, 3, 3, 3, 5, 5, 5, 8, 8, 8, 9, 9, 9, 12, 12, 13, 13, 17, 17, 17, 19, 19, 19, 25, 25, 25, 25, 25, 25, 25, 25, 29, 29, 29, 29, 32, 32, 33, 33, 35, 35, 41, 41, 44, 45, 49, 49];
for (let i = 0; i < values.length; i++) {
  chart.ChartData._get_Item(i + 1, 1).NumberValue = values[i];
}

//Set series label
chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 0, 0, 0);

//Set values for series
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemRCLL(1, 0, values.length, 0);

chart.PrimaryCategoryAxis.NumberOfBins = 7;
chart.PrimaryCategoryAxis.GapWidth = 20;
//Chart title
chart.ChartTitle.TextProperties.Text = "Histogram";
chart.ChartLegend.Position = wasmModule.ChartLegendPositionType.Bottom;
```

---

# Spire.Presentation JavaScript Chart
## Create Line Markers Chart in PowerPoint
```javascript
//Create a PPT file
let ppt = wasmModule.Presentation.Create();

//Add line markers chart
let rect1 = wasmModule.RectangleF.FromLTRB(90, 100, 640, 420);
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.LineMarkers, rectangle: rect1, init: false });

//Chart title
chart.ChartTitle.TextProperties.Text = "Line Makers Chart";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Data for series
let Series1 = [7.7, 8.9, 1.0, 2.4];
let Series2 = [15.2, 5.3, 6.7, 8];

//Set series text
chart.ChartData._get_Item(0, 1).Text = "Series1";
chart.ChartData._get_Item(0, 2).Text = "Series2";

//Set category text
chart.ChartData._get_Item(1, 0).Text = "Category 1";
chart.ChartData._get_Item(2, 0).Text = "Category 2";
chart.ChartData._get_Item(3, 0).Text = "Category 3";
chart.ChartData._get_Item(4, 0).Text = "Category 4";

//Fill data for chart
for (let i = 0; i < Series1.length; ++i) {
  chart.ChartData._get_Item(i + 1, 1).NumberValue = Series1[i];
  chart.ChartData._get_Item(i + 1, 2).NumberValue = Series2[i];
}

//Set series label
chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("B1", "C1");
//Set category label
chart.Categories.CategoryLabels = chart.ChartData._get_ItemNE("A2", "A5");

//Set values for series
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemNE("B2", "B5");
chart.Series.get_Item(1).Values = chart.ChartData._get_ItemNE("C2", "C5");
```

---

# Spire.Presentation JavaScript Map Chart
## Create a map chart in a PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Insert a Map chart to the first slide
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Map, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 500, 500), init: false });
chart.ChartData._get_Item(0, 1).Text = "series";

//Define some data.
let countries = ["China", "Russia", "France", "Mexico", "United States", "India", "Australia"];
for (let i = 0; i < countries.length; i++) {
  chart.ChartData._get_Item(i + 1, 0).Text = countries[i];
}
let values = [32, 20, 23, 17, 18, 6, 11];
for (let i = 0; i < values.length; i++) {
  chart.ChartData._get_Item(i + 1, 1).NumberValue = values[i];
}
chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 1, 0, 1);
chart.Categories.CategoryLabels = chart.ChartData._get_ItemRCLL(1, 0, 7, 0);
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemRCLL(1, 1, 7, 1);
```

---

# spire.presentation javascript chart
## create Pareto chart in PowerPoint presentation
```javascript
//Create PPT document
let ppt = wasmModule.Presentation.Create();

//Create a Pareto chart in first slide
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Pareto, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 550, 450), init: false });

//Set series text
chart.ChartData._get_Item(0, 1).Text = "Series 1";

//Set chart data and labels
chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 1, 0, 1);
chart.Categories.CategoryLabels = chart.ChartData._get_ItemRCLL(1, 0, categories.length, 0);
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemRCLL(1, 1, values.length, 1);

//Configure chart appearance
chart.PrimaryCategoryAxis.IsBinningByCategory = true;
chart.Series.get_Item(1).Line.FillFormat.FillType = wasmModule.FillFormatType.Solid;
chart.Series.get_Item(1).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_Red();
chart.ChartTitle.TextProperties.Text = "Pareto";
chart.HasLegend = true;
chart.ChartLegend.Position = wasmModule.ChartLegendPositionType.Bottom;
```

---

# Create Pie Chart in PowerPoint
## This code demonstrates how to create a pie chart in a PowerPoint presentation using JavaScript.
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Insert a Pie chart to the first slide and set the chart title.
let rect1 = wasmModule.RectangleF.FromLTRB(40, 100, 590, 420);
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Pie, rectangle: rect1, init: false });
chart.ChartTitle.TextProperties.Text = "Sales by Quarter";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Define some data.
let quarters = ["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"];
let sales = [210, 320, 180, 500];

//Append data to ChartData, which represents a data table where the chart data is stored.
chart.ChartData._get_Item(0, 0).Text = "Quarters";
chart.ChartData._get_Item(0, 1).Text = "Sales";
for (let i = 0; i < quarters.length; ++i) {
  chart.ChartData._get_Item(i + 1, 0).Text = quarters[i];
  chart.ChartData._get_Item(i + 1, 1).NumberValue = sales[i];
}

//Set category labels, series label and series data.
chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("B1", "B1");
chart.Categories.CategoryLabels = chart.ChartData._get_ItemNE("A2", "A5");
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemNE("B2", "B5");

//Add data points to series and fill each data point with different color.
for (let i = 0; i < chart.Series.get_Item(0).Values.Count; i++) {
  let cdp = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
  cdp.Index = i;
  chart.Series.get_Item(0).DataPoints.Add(cdp);
}
chart.Series.get_Item(0).DataPoints.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
chart.Series.get_Item(0).DataPoints.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_RosyBrown();
chart.Series.get_Item(0).DataPoints.get_Item(1).Fill.FillType = wasmModule.FillFormatType.Solid;
chart.Series.get_Item(0).DataPoints.get_Item(1).Fill.SolidColor.Color = wasmModule.Color.get_LightBlue();
chart.Series.get_Item(0).DataPoints.get_Item(2).Fill.FillType = wasmModule.FillFormatType.Solid;
chart.Series.get_Item(0).DataPoints.get_Item(2).Fill.SolidColor.Color = wasmModule.Color.get_LightPink();
chart.Series.get_Item(0).DataPoints.get_Item(3).Fill.FillType = wasmModule.FillFormatType.Solid;
chart.Series.get_Item(0).DataPoints.get_Item(3).Fill.SolidColor.Color = wasmModule.Color.get_MediumPurple();

//Set the data labels to display label value and percentage value.
chart.Series.get_Item(0).DataLabels.LabelValueVisible = true;
chart.Series.get_Item(0).DataLabels.PercentValueVisible = true;
```

---

# spire presentation javascript scatter chart
## create scatter chart with markers
```javascript
//Creat a presentation
let ppt = wasmModule.Presentation.Create();

//Set background image
let rect2 = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: inputFileName, rectangle: rect2 });
ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

//Insert a chart and set chart title and chart type
let rect1 = wasmModule.RectangleF.FromLTRB(90, 100, 640, 420);
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.ScatterMarkers, rectangle: rect1, init: false });
chart.ChartTitle.TextProperties.Text = "ScatterMarker Chart";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Set chart data
let xdata = [2.7, 8.9, 10.0, 12.4];
let ydata = [3.2, 15.3, 6.7, 8];

chart.ChartData._get_Item(0, 0).Text = "X-Value";
chart.ChartData._get_Item(0, 1).Text = "Y-Value";

for (let i = 0; i < xdata.length; ++i) {
  chart.ChartData._get_Item(i + 1, 0).NumberValue = xdata[i];
  chart.ChartData._get_Item(i + 1, 1).NumberValue = ydata[i];
}

//Set the series label
chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("B1", "B1");

//Assign data to X axis, Y axis and Bubbles
chart.Series.get_Item(0).XValues = chart.ChartData._get_ItemNE("A2", "A5");
chart.Series.get_Item(0).YValues = chart.ChartData._get_ItemNE("B2", "B5");
```

---

# Spire.Presentation JavaScript Sunburst Chart
## Create a Sunburst chart in a PowerPoint presentation
```javascript
//Create PPT document
let ppt = wasmModule.Presentation.Create();

//Create a SunBurst chart to the first slide
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.SunBurst, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 550, 450), init: false });

//Set series text
chart.ChartData._get_Item(0, 3).Text = "Series 1";

//Set category text
let categories = [["Branch 1", "Stem 1", "Leaf 1"], ["Branch 1", "Stem 1", "Leaf 2"], ["Branch 1", "Stem 1", "Leaf 3"],
["Branch 1", "Stem 2", "Leaf 4"], ["Branch 1", "Stem 2", "Leaf 5"], ["Branch 1", "Leaf 6", ""], ["Branch 1", "Leaf 7", ""],
["Branch 2", "Stem 3", "Leaf 8"], ["Branch 2", "Leaf 9", ""], ["Branch 2", "Stem 4", "Leaf 10"], ["Branch 2", "Stem 4", "Leaf 11"],
["Branch 2", "Stem 5", "Leaf 12"], ["Branch 3", "Stem 5", "Leaf 13"], ["Branch 3", "Stem 6", "Leaf 14"], ["Branch 3", "Leaf 15", ""]];

for (let i = 0; i < categories.length; i++) {
  for (let j = 0; j < categories[0].length; j++) {
    chart.ChartData._get_Item(i + 1, j).Text = categories[i][j];
  }
}
//Fill data for chart
let values = [17, 23, 48, 22, 76, 54, 77, 26, 44, 63, 10, 15, 48, 15, 51];
for (let i = 0; i < values.length; i++) {
  chart.ChartData._get_Item(i + 1, 3).NumberValue = values[i];
}

//Set series labels
chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 3, 0, 3);

//Set categories labels
chart.Categories.CategoryLabels = chart.ChartData._get_ItemRCLL(1, 0, values.length, 2);

//Assign data to series values
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemRCLL(1, 3, values.length, 3);

chart.Series.get_Item(0).DataLabels.CategoryNameVisible = true;
chart.ChartTitle.TextProperties.Text = "SunBurst";
chart.HasLegend = true;
chart.ChartLegend.Position = wasmModule.ChartLegendPositionType.Top;
```

---

# Spire.Presentation JavaScript TreeMap Chart
## Create a TreeMap chart in a PowerPoint presentation
```javascript
//Create PPT document
let ppt = wasmModule.Presentation.Create();

//Create a TreeMap chart to the first slide
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.TreeMap, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 550, 450), init: false });

//Set series text
chart.ChartData._get_Item(0, 3).Text = "Series 1";

//Set category text
let categories = [["Branch 1", "Stem 1", "Leaf 1"], ["Branch 1", "Stem 1", "Leaf 2"], ["Branch 1", "Stem 1", "Leaf 3"],
["Branch 1", "Stem 2", "Leaf 4"], ["Branch 1", "Stem 2", "Leaf 5"], ["Branch 1", "Stem 2", "Leaf 6"], ["Branch 1", "Stem 2", "Leaf 7"],
["Branch 2", "Stem 3", "Leaf 8"], ["Branch 2", "Stem 3", "Leaf 9"], ["Branch 2", "Stem 4", "Leaf 10"], ["Branch 2", "Stem 4", "Leaf 11"],
["Branch 2", "Stem 5", "Leaf 12"], ["Branch 3", "Stem 5", "Leaf 13"], ["Branch 3", "Stem 6", "Leaf 14"], ["Branch 3", "Stem 6", "Leaf 15"]];
for (let i = 0; i < categories.length; i++) {
  for (let j = 0; j < categories[0].length; j++) {
    chart.ChartData._get_Item(i + 1, j).Text = categories[i][j];
  }
}

//Fill data for chart
let values = [17, 23, 48, 22, 76, 54, 77, 26, 44, 63, 10, 15, 48, 15, 51];
for (let i = 0; i < values.length; i++) {
  chart.ChartData._get_Item(i + 1, 3).NumberValue = values[i];
}

//Set series labels
chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 3, 0, 3);

//Set categories labels
chart.Categories.CategoryLabels = chart.ChartData._get_ItemRCLL(1, 0, values.length, 2);

//Assign data to series values
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemRCLL(1, 3, values.length, 3);

chart.Series.get_Item(0).DataLabels.CategoryNameVisible = true;
chart.Series.get_Item(0).TreeMapLabelOption = wasmModule.TreeMapLabelOption.Banner;
chart.ChartTitle.TextProperties.Text = "TreeMap";
chart.HasLegend = true;
chart.ChartLegend.Position = wasmModule.ChartLegendPositionType.Top;
```

---

# spire.presentation javascript waterfall chart
## create WaterFall chart in presentation
```javascript
//Create PPT document
let ppt = wasmModule.Presentation.Create();

//Create a WaterFall chart to the first slide
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.WaterFall, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 550, 450), init: false });

//Set series text
chart.ChartData._get_Item(0, 1).Text = "Series 1";

//Set category text
let categories = ["Category 1", "Category 2", "Category 3", "Category 4", "Category 5", "Category 6", "Category 7"];
for (let i = 0; i < categories.length; i++) {
  chart.ChartData._get_Item(i + 1, 0).Text = categories[i];
}

//Fill data for chart
let values = [100, 20, 50, -40, 130, -60, 70];
for (let i = 0; i < values.length; i++) {
  chart.ChartData._get_Item(i + 1, 1).NumberValue = values[i];
}

//Set series labels
chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 1, 0, 1);

//Set categories labels
chart.Categories.CategoryLabels = chart.ChartData._get_ItemRCLL(1, 0, categories.length, 0);

//Assign data to series values
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemRCLL(1, 1, values.length, 1);

//Operate the third datapoint of first series
let chartDataPoint = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
chartDataPoint.Index = 2;
chartDataPoint.SetAsTotal = true;
chart.Series.get_Item(0).DataPoints.Add(chartDataPoint);

//Operate the sixth datapoint of first series
let chartDataPoint2 = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
chartDataPoint2.Index = 5;
chartDataPoint2.SetAsTotal = true;
chart.Series.get_Item(0).DataPoints.Add(chartDataPoint2);
chart.Series.get_Item(0).ShowConnectorLines = true;
chart.Series.get_Item(0).DataLabels.LabelValueVisible = true;

chart.ChartLegend.Position = wasmModule.ChartLegendPositionType.Right;
chart.ChartTitle.TextProperties.Text = "WaterFall";
```

---

# PowerPoint Chart Legend Management
## Delete chart legend entries from PowerPoint presentation
```javascript
// Create a PowerPoint document
let ppt = wasmModule.Presentation.Create();

// Load the file from disk
ppt.LoadFromFile(inputFileName);

// Get the chart
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Delete the first and the second legend entries from the chart
chart.ChartLegend.DeleteEntry(0);
chart.ChartLegend.DeleteEntry(1);
```

---

# Chart Switch Row and Column Detection
## Detect whether a chart has the "SwitchRowAndColumn" setting
```javascript
// Create a PowerPoint document
let ppt = wasmModule.Presentation.Create();

// Load the file
ppt.LoadFromFile(inputFileName);

// Get the chart
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

let stringBuilder = [];

// Detect whether the chart has "SwitchRowAndColumn" setting
let result = chart.IsSwitchRowAndColumn();
stringBuilder.push("'SwitchRowAndColumn' value of the chart is " + result + "\n");
```

---

# Spire.Presentation JavaScript Chart
## Set doughnut chart hole size
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

// Get the chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Set hole size
Chart.Series.get_Item(0).DoughnutHoleSize = 55;
```

---

# Edit Chart Data in PowerPoint
## Modify data point values in a chart within a PowerPoint presentation
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

// Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Change the value of the second datapoint of the first series
Chart.Series.get_Item(0).Values.get_Item(1).NumberValue = 6;
```

---

# spire.presentation javascript chart
## explode pie chart
```javascript
//Get the chart that needs to set the point explosion.
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

chart.Series.get_Item(0).Distance = 15;
```

---

# Spire.Presentation JavaScript Chart Marker
## Fill picture in chart marker
```javascript
//Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

//Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Load image file in ppt
let stream = wasmModule.Stream.CreateByFile(inputFileNameImg);
let IImage = ppt.Images.Append({ stream: stream });

//Create a ChartDataPoint object and specify the index
let dataPoint = wasmModule.ChartDataPoint.Create(Chart.Series.get_Item(0));
dataPoint.Index = 0;

//Fill picture in marker
dataPoint.MarkerFill.Fill.FillType = wasmModule.FillFormatType.Picture;
dataPoint.MarkerFill.Fill.PictureFill.Picture.EmbedImage = IImage;

//Set marker size
dataPoint.MarkerSize = 20;

//Add the data point in series
Chart.Series.get_Item(0).DataPoints.Add(dataPoint);
```

---

# spire presentation javascript chart
## format chart data labels
```javascript
//Get the chart
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Get the chart series
let sers = chart.Series;

//Initialize four instances of series label and set parameters of each label
let cd1 = sers.get_Item(0).DataLabels.Add();
cd1.PercentageVisible = true;
cd1.TextFrame.Text = "Custom Datalabel1";
cd1.TextFrame.TextRange.FontHeight = 12;
cd1.TextFrame.TextRange.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");
cd1.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
cd1.TextFrame.TextRange.Fill.SolidColor.Color = wasmModule.Color.get_Green();

let cd2 = sers.get_Item(0).DataLabels.Add();
cd2.Position = wasmModule.ChartDataLabelPosition.InsideEnd;
cd2.PercentageVisible = true;
cd2.TextFrame.Text = "Custom Datalabel2";
cd2.TextFrame.TextRange.FontHeight = 10;
cd2.TextFrame.TextRange.LatinFont = wasmModule.TextFont.Create("Arial");
cd2.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
cd2.TextFrame.TextRange.Fill.SolidColor.Color = wasmModule.Color.get_OrangeRed();

let cd3 = sers.get_Item(0).DataLabels.Add();
cd3.Position = wasmModule.ChartDataLabelPosition.Center;
cd3.PercentageVisible = true;
cd3.TextFrame.Text = "Custom Datalabel3";
cd3.TextFrame.TextRange.FontHeight = 14;
cd3.TextFrame.TextRange.LatinFont = wasmModule.TextFont.Create("Calibri");
cd3.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
cd3.TextFrame.TextRange.Fill.SolidColor.Color = wasmModule.Color.get_Blue();

let cd4 = sers.get_Item(0).DataLabels.Add();
cd4.Position = wasmModule.ChartDataLabelPosition.InsideBase;
cd4.PercentageVisible = true;
cd4.TextFrame.Text = "Custom Datalabel4";
cd4.TextFrame.TextRange.FontHeight = 12;
cd4.TextFrame.TextRange.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");
cd4.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
cd4.TextFrame.TextRange.Fill.SolidColor.Color = wasmModule.Color.get_OliveDrab();
```

---

# spire.presentation javascript chart
## get values and unit from chart axis
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

// Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Get unit from primary category axis
let MajorUnit = Chart.PrimaryCategoryAxis.MajorUnit;
let type = Chart.PrimaryCategoryAxis.MajorUnitScale;

// Get values from primary value axis
let minValue = Chart.PrimaryValueAxis.MinValue;
let maxValue = Chart.PrimaryValueAxis.MaxValue;
```

---

# spire presentation javascript chart
## group two-level axis labels in chart
```javascript
//Create a PowerPoint document.
let ppt = wasmModule.Presentation.Create();

//Load the file from disk.
ppt.LoadFromFile(inputFileName);

//Get the chart.
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Get the category axis from the chart.
let chartAxis = chart.PrimaryCategoryAxis;

//Group the axis labels that have the same first-level label.
if (chartAxis.HasMultiLvlLbl) {
  chartAxis.IsMergeSameLabel = true;
}
```

---

# Hide Chart Axis and Gridlines in PowerPoint
## This code demonstrates how to hide chart axis and gridlines in a PowerPoint presentation using JavaScript
```javascript
//Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Hide axis
Chart.PrimaryCategoryAxis.IsVisible = false;
Chart.PrimaryValueAxis.IsVisible = false;

//Remove gridline
Chart.PrimaryValueAxis.MajorGridTextLines.FillType = wasmModule.FillFormatType.None;
```

---

# Hide or Show Chart Series
## Demonstrate how to hide or show a series in a chart within a PowerPoint presentation
```javascript
//Get the first slide.
let slide = ppt.Slides.get_Item(0);

//Get the first chart.
let chart = slide.Shapes.get_Item(0);

//Hide the first series of the chart.
chart.Series.get_Item(0).IsHidden = true;

//Show the first series of the chart.
//chart.Series[0].IsHidden = false;
```

---

# spire presentation javascript invert if negative
## set invert if negative property for chart series
```javascript
//Get chart from the first slide
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Set invert if negative for the first series
chart.Series.get_Item(0).InvertIfNegative = true;
```

---

# Spire.Presentation JavaScript Chart
## Modify chart category axis settings
```javascript
//Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Modify the major unit
Chart.PrimaryCategoryAxis.IsAutoMajor = false;
Chart.PrimaryCategoryAxis.MajorUnit = 1;
Chart.PrimaryCategoryAxis.MajorUnitScale = wasmModule.ChartBaseUnitType.Months;
```

---

# Spire.Presentation JavaScript Chart
## Create Multi-Category Chart in PowerPoint
```javascript
//Create a PPT file
let ppt = wasmModule.Presentation.Create();

//Add line markers chart
let rect1 = wasmModule.RectangleF.FromLTRB(90, 100, 640, 420);
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.ColumnClustered, rectangle: rect1, init: false });

//Chart title
chart.ChartTitle.TextProperties.Text = "Muli-Category";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Data for series
let Series1 = [7.7, 8.9, 7, 6, 7, 8];

//Set series text
chart.ChartData._get_Item(0, 2).Text = "Series1";

//Set category text
chart.ChartData._get_Item(1, 0).Text = "Grp 1";
chart.ChartData._get_Item(3, 0).Text = "Grp 2";
chart.ChartData._get_Item(5, 0).Text = "Grp 3";

chart.ChartData._get_Item(1, 1).Text = "A";
chart.ChartData._get_Item(2, 1).Text = "B";
chart.ChartData._get_Item(3, 1).Text = "C";
chart.ChartData._get_Item(4, 1).Text = "D";
chart.ChartData._get_Item(5, 1).Text = "E";
chart.ChartData._get_Item(6, 1).Text = "F";

//Fill data for chart
for (let i = 0; i < Series1.length; ++i) {
  chart.ChartData._get_Item(i + 1, 2).NumberValue = Series1[i];
}

//Set series label
chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("C1", "C1");
//Set category label
chart.Categories.CategoryLabels = chart.ChartData._get_ItemNE("A2", "B7");

//Set values for series
chart.Series.get_Item(0).Values = chart.ChartData._get_ItemNE("C2", "C7");

//Set if the category axis has multiple levels
chart.PrimaryCategoryAxis.HasMultiLvlLbl = true;
//Merge same label
chart.PrimaryCategoryAxis.IsMergeSameLabel = true;
```

---

# spire.presentation javascript chart protection
## protect chart in PowerPoint document
```javascript
//Get the first shape from slide and convert it as IChart.
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Set the Boolean value of IChart.IsDataProtect as true.
chart.IsDataProtect = true;
```

---

# Remove Chart from PowerPoint Slide
## Remove chart shapes from a PowerPoint slide
```javascript
//Create a PowerPoint document
let ppt = wasmModule.Presentation.Create();

//Load the file from disk.
ppt.LoadFromFile(inputFileName);

//Get the first slide from the document.
let slide = ppt.Slides.get_Item(0);

//Remove chart from the slide.
for (let i = 0; i < slide.Shapes.Count; i++) {
  let shape = slide.Shapes.get_Item(i);
  if (shape instanceof wasmModule.IChart) {
    slide.Shapes.Remove(shape);
  }
}
```

---

# spire.presentation javascript chart
## remove tick marks of axis
```javascript
//Get the chart that need to be adjusted the number format and remove the tick marks.
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Set percentage number format for the axis value of chart.
chart.PrimaryValueAxis.NumberFormat = "0#\\%";

//Remove the tick marks for value axis and category axis.
chart.PrimaryValueAxis.MajorTickMark = wasmModule.TickMarkType.TickMarkNone;
chart.PrimaryValueAxis.MinorTickMark = wasmModule.TickMarkType.TickMarkNone;
chart.PrimaryCategoryAxis.MajorTickMark = wasmModule.TickMarkType.TickMarkNone;
chart.PrimaryCategoryAxis.MinorTickMark = wasmModule.TickMarkType.TickMarkNone;
```

---

# spire.presentation javascript chart
## save chart as image
```javascript
let ppt = wasmModule.Presentation.Create();
//Load PPT file from disk
ppt.LoadFromFile(inputFileName);

//Save chart as image in .png format
let image = ppt.Slides.get_Item(0).Shapes.SaveAsImage({ shapeIndex: 0 });

// Define the output file name
const outputFileName = "SaveChartAsImage_out.png";

// Save the document to the specified path
image.Save(outputFileName);

// Clean up resources
image.Dispose();
ppt.Dispose();
```

---

# Scale Bubble Chart Size in Presentation
## Adjust bubble scale in a PowerPoint bubble chart
```javascript
//Get the chart from the first presentation slide.
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Scale the bubble size, the range value is from 0 to 300.
chart.BubbleScale = 50;
```

---

# spire presentation javascript chart
## set chart axis position
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

// Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Set axis position
Chart.PrimaryValueAxis.CrossBetweenType = wasmModule.CrossBetweenType.MidpointOfCategory;
```

---

# Setting Chart Axis Type in PowerPoint
## Set primary category axis type to date axis with month scale
```javascript
//Get the chart
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(1);

chart.PrimaryCategoryAxis.AxisType = wasmModule.AxisType.DateAxis;
chart.PrimaryCategoryAxis.MajorUnitScale = wasmModule.ChartBaseUnitType.Months;
```

---

# spire presentation chart border style
## set chart border style in presentation
```javascript
//Get chart on the first slide
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Set border style
chart.Line.FillFormat.FillType = wasmModule.FillFormatType.Solid;
chart.Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_Red();
chart.BorderRoundedCorners = true;
```

---

# Set Chart Data Label Range in PowerPoint
## Configure data labels for charts in PowerPoint presentations
```javascript
// Create a PowerPoint document
let ppt = wasmModule.Presentation.Create();

// Add a ColumnStacked chart
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ 
  type: wasmModule.ChartType.ColumnStacked, 
  rectangle: wasmModule.RectangleF.FromLTRB(100, 100, 600, 500) 
});

// Set data for the chart
let cellRange = chart.ChartData._get_ItemN("F1");
cellRange.Text = "labelA";
cellRange = chart.ChartData._get_ItemN("F2");
cellRange.Text = "labelB";
cellRange = chart.ChartData._get_ItemN("F3");
cellRange.Text = "labelC";
cellRange = chart.ChartData._get_ItemN("F4");
cellRange.Text = "labelD";

// Set data label ranges
chart.Series.get_Item(0).DataLabelRanges = chart.ChartData._get_ItemNE("F1", "F4");

// Add data label
let dataLabel1 = chart.Series.get_Item(0).DataLabels.Add();
dataLabel1.ID = 0;
// Show the value
dataLabel1.LabelValueVisible = false;
// Show the label string
dataLabel1.ShowDataLabelsRange = true;
```

---

# spire.presentation javascript chart number format
## Set number format for chart data in PowerPoint presentation
```javascript
//Get chart on the first slide
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Set the number format for Axis
chart.PrimaryValueAxis.NumberFormat = "#,##0.00";

//Set the DataLabels format for Axis
chart.Series.get_Item(0).DataLabels.LabelValueVisible = true;
chart.Series.get_Item(0).DataLabels.PercentValueVisible = false;
chart.Series.get_Item(0).DataLabels.NumberFormat = "#,##0.00";
chart.Series.get_Item(0).DataLabels.HasDataSource = false;

//Set the number format for ChartData
for (let i = 1; i <= chart.Series.get_Item(0).Values.Count; i++) {
  chart.ChartData._get_Item(i, 1).NumberFormat = "#,##0.00";
}
```

---

# spire presentation javascript trendline
## set color and name for trendline in chart
```javascript
//Find the first trendline in the chart
let trendline = chart.Series.get_Item(0).TrendLines[0];

//Set name for trendline
trendline.Name = "trendlineName";

//Set color for trendline
trendline.Line.FillType = wasmModule.FillFormatType.Solid;
trendline.Line.SolidFillColor.Color = wasmModule.Color.get_Red();
```

---

# spire.presentation javascript chart
## set data label position in chart
```javascript
//Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

//Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Add data label
let label = Chart.Series.get_Item(0).DataLabels.Add();
//Set the position of the label
label.X = 0.1;
label.Y = 0.1;

// Save the document to the specified path
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });
```

---

# spire presentation javascript chart
## set datapoint color in chart
```javascript
//Get the chart
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Initialize an instances of dataPoint
let cdp1 = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));

//Specify the datapoint order
cdp1.Index = 0;

//Set the color of the datapoint
cdp1.Fill.FillType = wasmModule.FillFormatType.Solid;
cdp1.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Orange;

//Add the dataPoint to first series
chart.Series.get_Item(0).DataPoints.Add(cdp1);

//Set the color for the other three data points
let cdp2 = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
cdp2.Index = 1;
cdp2.Fill.FillType = wasmModule.FillFormatType.Solid;
cdp2.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Gold;
chart.Series.get_Item(0).DataPoints.Add(cdp2);

let cdp3 = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
cdp3.Index = 2;
cdp3.Fill.FillType = wasmModule.FillFormatType.Solid;
cdp3.Fill.SolidColor.KnownColor = wasmModule.KnownColors.MediumPurple;
chart.Series.get_Item(0).DataPoints.Add(cdp3);

let cdp4 = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
cdp4.Index = 1;
cdp4.Fill.FillType = wasmModule.FillFormatType.Solid;
cdp4.Fill.SolidColor.KnownColor = wasmModule.KnownColors.ForestGreen;
chart.Series.get_Item(0).DataPoints.Add(cdp4);
```

---

# spire presentation javascript chart
## set display unit of value axis for chart
```javascript
//Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Set the display unit
Chart.PrimaryValueAxis.DisplayUnit = wasmModule.ChartDisplayUnitType.Hundreds;
```

---

# spire.presentation javascript chart
## set distance from axis for chart in PowerPoint
```javascript
// Create a ppt document
let ppt = wasmModule.Presentation.Create();

// Append ColumnClustered chart
let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.ColumnClustered, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 450, 450) });

// Get the PrimaryCategory axis
let chartAxis = chart.PrimaryCategoryAxis;

// Set "Distance from axis"
chartAxis.LabelsDistance = 200;
```

---

# spire.presentation javascript chart
## set gap width for chart in powerpoint
```javascript
// Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Set gap width
Chart.GapWidth = 50;
```

---

# spire presentation javascript chart
## set legend options for chart in powerpoint
```javascript
// Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Set the legend position
Chart.ChartLegend.Left = 20;
Chart.ChartLegend.Top = 20;

// Set the legend size
Chart.ChartLegend.Width = 250;
Chart.ChartLegend.Height = 30;
```

---

# spire.presentation javascript chart
## set number format for category axis
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile("ChartSample3.pptx");

// Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Set the number format
Chart.PrimaryCategoryAxis.NumberFormat = "yyyy";
```

---

# spire presentation javascript chart
## set percentage for labels in stacked column chart
```javascript
// Get the chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

let dataPontPercent = 0;

for (let i = 0; i < Chart.Series.Count; i++) {
  let series = Chart.Series.get_Item(i);
  //Get the total number
  let total = GetTotal(series.Values);
  for (let j = 0; j < series.Values.Count; j++) {
    //Get the percent
    dataPontPercent = parseFloat(series.Values.get_Item(j).Text) / total * 100;
    //Add datalabels
    let label = series.DataLabels.Add();
    label.LabelValueVisible = true;
    //Set the percent text for the label
    label.TextFrame.Paragraphs.get_Item(0).Text = `${dataPontPercent.toFixed(2)} %`;
    label.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).FontHeight = 12;
  }
}

function GetTotal(ranges) {
  let total = 0;
  for (let i = 0; i < ranges.Count; i++) {
    total += parseFloat(ranges.get_Item(i).Text);
  }
  return total;
}
```

---

# Spire.Presentation JavaScript Chart Data Labels
## Set Position and Properties of Chart Data Labels
```javascript
//Get the chart.
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Add data label to chart and set its id.
let label1 = chart.Series.get_Item(0).DataLabels.Add();
label1.ID = 0;

// Set the default position of data label. This position is relative to the data markers.
//label1.Position = ChartDataLabelPosition.OutsideEnd;

// Set custom position of data label. This position is relative to the default position.
label1.X = 0.1;
label1.Y = -0.1;

// Set label value visible
label1.LabelValueVisible = true;

// Set legend key invisible
label1.LegendKeyVisible = false;

// Set category name invisible
label1.CategoryNameVisible = false;

// Set series name invisible
label1.SeriesNameVisible = false;

// Set Percentage invisible
label1.PercentageVisible = false;

// Set border style and fill style of data label
label1.Line.FillType = wasmModule.FillFormatType.Solid;
label1.Line.SolidFillColor.Color = wasmModule.Color.get_Blue();
label1.Fill.FillType = wasmModule.FillFormatType.Solid;
label1.Fill.SolidColor.Color = wasmModule.Color.get_Orange();
```

---

# spire.presentation javascript chart title
## set rotation angle for chart title
```javascript
// Create a PowerPoint document 
let ppt = wasmModule.Presentation.Create();

// Load file from VFS
ppt.LoadFromFile(inputFileName);

// Get the chart
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);
// Set rotation angle
chart.ChartTitle.TextProperties.RotationAngle = -30;
```

---

# spire.presentation javascript chart
## set rotation angle for data labels
```javascript
// Get the first chart
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Set the rotation angle for the datalabels of first serie
for (let i = 0; i < Chart.Series.get_Item(0).Values.Count; i++) {
  let datalabel = Chart.Series.get_Item(0).DataLabels.Add();
  datalabel.ID = i;
  datalabel.RotationAngle = 45;
}
```

---

# Spire.Presentation JavaScript Chart Text Rotation
## Set rotation angle for value axis text in presentation chart
```javascript
//Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Set the rotation angle for the text on the value axis
Chart.PrimaryValueAxis.TextRotationAngle = 45;
```

---

# PowerPoint Chart Series Line Color
## Set series line color for chart in PowerPoint presentation
```javascript
// Create a PowerPoint document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

// Get the first chart
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(0);
if (shape instanceof wasmModule.IChart) {
  let chart = shape;
  let seriesLine = chart.SeriesLine;
  seriesLine.FillType = wasmModule.FillFormatType.Solid;

  // Set the color of seriesLine
  seriesLine.FillFormat.SolidFillColor.Color = wasmModule.Color.get_Red();
}
```

---

# Spire.Presentation JavaScript Chart Series Overlap
## Set series overlap for chart in PowerPoint presentation
```javascript
// Create PPT document
let ppt = wasmModule.Presentation.Create();

// Get chart on the first slide
let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Set overlap
Chart.OverLap = 50;
```

---

# spire presentation javascript chart marker
## set size and style for data marker
```javascript
// Get the first chart
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

for (let i = 0; i < chart.Series.get_Item(0).Values.Count; i++) {
  // Create a ChartDataPoint object and specify the index.
  let dataPoint = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
  dataPoint.Index = i;

  // Set the fill color of the data marker.
  dataPoint.MarkerFill.Fill.FillType = wasmModule.FillFormatType.Solid;
  dataPoint.MarkerFill.Fill.SolidColor.Color = wasmModule.Color.get_Yellow();

  // Set the line color of the data marker.
  dataPoint.MarkerFill.Line.FillType = wasmModule.FillFormatType.Solid;
  dataPoint.MarkerFill.Line.SolidFillColor.Color = wasmModule.Color.get_YellowGreen();

  // Set the size of the data marker.
  dataPoint.MarkerSize = 20;

  // Set the style of the data marker
  dataPoint.MarkerStyle = wasmModule.ChartMarkerType.Diamond;
  chart.Series.get_Item(0).DataPoints.Add(dataPoint);
}
```

---

# Spire.Presentation JavaScript Chart
## Set size for chart plot area
```javascript
// Get chart on the first chart
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Set width and height for chart plot area
chart.PlotArea.Width = 250;
chart.PlotArea.Height = 300;
```

---

# spire presentation javascript chart
## set text font for chart title
```javascript
//Get the chart.
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Set the font for the text on chart title area
chart.ChartTitle.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.LatinFont = wasmModule.TextFont.Create("Arial Unicode MS");
chart.ChartTitle.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Blue;
chart.ChartTitle.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.FontHeight = 50;
```

---

# spire presentation chart text formatting
## set font for chart legend and axis text
```javascript
// Get the chart
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Set the font for the text on Chart Legend area
chart.ChartLegend.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Green;
chart.ChartLegend.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.LatinFont = wasmModule.TextFont.Create("Arial Unicode MS");

// Set the font for the text on Chart Axis area
chart.PrimaryCategoryAxis.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Red;
chart.PrimaryCategoryAxis.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.Fill.FillType = wasmModule.FillFormatType.Solid;
chart.PrimaryCategoryAxis.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.FontHeight = 10;
chart.PrimaryCategoryAxis.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.LatinFont = wasmModule.TextFont.Create("Arial Unicode MS");
```

---

# Spire.Presentation JavaScript Chart
## Set tick-mark labels on category axis
```javascript
// Get the chart from the PowerPoint slide
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Rotate tick labels
chart.PrimaryCategoryAxis.TextRotationAngle = 45;

// Specify interval between labels
chart.PrimaryCategoryAxis.IsAutomaticTickLabelSpacing = false;
chart.PrimaryCategoryAxis.TickLabelSpacing = 2;

// Change position
chart.PrimaryCategoryAxis.TickLabelPosition = wasmModule.TickLabelPositionType.TickLabelPositionHigh;
```

---

# spire presentation javascript chart
## set tick marks interval for chart
```javascript
// Get the first chart
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);
let chartAxis = chart.PrimaryCategoryAxis;
chartAxis.TickMarkSpacing = 2;
```

---

# spire presentation javascript chart labels
## show data labels in chart
```javascript
//Get chart on the first slide
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Show data labels
chart.Series.get_Item(0).DataLabels.LabelValueVisible = true;
chart.Series.get_Item(0).DataLabels.CategoryNameVisible = true;
chart.Series.get_Item(0).DataLabels.SeriesNameVisible = true;
```

---

# spire presentation javascript chart
## vary colors of data markers in same chart series
```javascript
// Get the chart from the presentation.
let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Create a ChartDataPoint object and specify the index.
let dataPoint = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
dataPoint.Index = 0;

// Set the fill color of the data marker.
dataPoint.MarkerFill.Fill.FillType = wasmModule.FillFormatType.Solid;
dataPoint.MarkerFill.Fill.SolidColor.Color = wasmModule.Color.get_Red();

// Set the line color of the data marker.
dataPoint.MarkerFill.Line.FillType = wasmModule.FillFormatType.Solid;
dataPoint.MarkerFill.Line.SolidFillColor.Color = wasmModule.Color.get_Red();

// Add the data point to the point collection of a series.
chart.Series.get_Item(0).DataPoints.Add(dataPoint);

dataPoint = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
dataPoint.Index = 1;
dataPoint.MarkerFill.Fill.FillType = wasmModule.FillFormatType.Solid;
dataPoint.MarkerFill.Fill.SolidColor.Color = wasmModule.Color.get_Black();
dataPoint.MarkerFill.Line.FillType = wasmModule.FillFormatType.Solid;
dataPoint.MarkerFill.Line.SolidFillColor.Color = wasmModule.Color.get_Black();
chart.Series.get_Item(0).DataPoints.Add(dataPoint);

dataPoint = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
dataPoint.Index = 2;
dataPoint.MarkerFill.Fill.FillType = wasmModule.FillFormatType.Solid;
dataPoint.MarkerFill.Fill.SolidColor.Color = wasmModule.Color.get_Blue();
dataPoint.MarkerFill.Line.FillType = wasmModule.FillFormatType.Solid;
dataPoint.MarkerFill.Line.SolidFillColor.Color = wasmModule.Color.get_Blue();
chart.Series.get_Item(0).DataPoints.Add(dataPoint);
```

---

# spire presentation javascript conversion
## convert ODP to PDF
```javascript
// Create PPT document
let ppt = wasmModule.Presentation.Create();

// Load the PPT document from VFS
ppt.LoadFromFile({ file: inputFileName, fileFormat: wasmModule.FileFormat.ODP });

// Define the output file name
const outputFileName = "ConvertODPtoPDF.pdf";

// Save the document 
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.PDF });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Conversion
## Convert PowerPoint to PDF with default font
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

// The font is preferred to convert to pdf or pictures, when the font used in the document is not installed in the system
wasmModule.Presentation.SetDefaultFontName("arial");

// Define the output file name
const outputFileName = "ConvertPdfWithDefaultFont.pdf";

// Save the document 
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.PDF });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Conversion
## Convert PPS to PPTX format
```javascript
// Create an instance of presentation document
let ppt = wasmModule.Presentation.Create();

// Load file
ppt.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ConvertPPSToPPTX.pptx";

// Save the document 
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Conversion
## Convert PowerPoint presentation to OFD format
```javascript
// Create an instance of presentation document
let ppt = wasmModule.Presentation.Create();

// Load file
ppt.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ConvertPPTToOFD.ofd";

// Save the document 
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.OFD });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Conversion
## Convert unhidden slides to PDF
```javascript
// Create an instance of presentation document
let ppt = wasmModule.Presentation.Create();

// Load PPT file from VFS
ppt.LoadFromFile(inputFileName);

// Convert the PPT unhidden slides to PDF file format
ppt.SaveToPdfOption.ContainHiddenSlides = false;

// Define the output file name
const outputFileName = "ConvertUnhideSlidesToPdf.pdf";

// Save the document 
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.PDF });
```

---

# spire.presentation javascript conversion
## convert PowerPoint to TIFF with custom size
```javascript
// Create PPT document
let ppt = wasmModule.Presentation.Create();

// Load the original PPT document from VFS
ppt.LoadFromFile(inputFileName);

// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Create a new PPT document
let newPpt = wasmModule.Presentation.Create();

// Remove the default slide
newPpt.Slides.RemoveAt(0);

// Define a new size
let size = wasmModule.SizeF.CreateWH(200, 200);

// Set PPT slide size
newPpt.SlideSize.Size = size;

// Insert the slide of original PPT
newPpt.Slides.Insert({ index: 0, slide: slide });

// Save the document 
newPpt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Tiff });

// Clean up resources
ppt.Dispose();
```

---

# spire.presentation javascript conversion
## convert individual slide to html
```javascript
// Create PPT document
let ppt = wasmModule.Presentation.Create();

// Load the PPT document from VFS
ppt.LoadFromFile(inputFileName);

// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Define the output file name
const outputFileName = "IndividualSlideToHtml.html";

// Save the document 
slide.SaveToFile(outputFileName, wasmModule.FileFormat.Html);

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Document Conversion
## Load and save DPS and DPT format documents
```javascript
// Load the input file into the virtual file system (VFS)
const inputFileName = "sample.dps";
await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

// Create PPT document
let ppt = wasmModule.Presentation.Create();

// Load the PPT document from VFS
ppt.LoadFromFile({ file: inputFileName, fileFormat:wasmModule.FileFormat.Dps });

// Define the output file name
const outputFileName = "LoadSaveDPSAndDPT.dps";

// Save the document 
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Dps });

// Clean up resources
ppt.Dispose();
```

---

# ODP to PDF Conversion
## Convert OpenDocument Presentation (ODP) files to PDF format
```javascript
// Define file names
const inputFileName = "toPdf.odp";
const outputFileName = "OdpToPdf.pdf";

// Create PPT document
let ppt = wasmModule.Presentation.Create();
// Load ODP file
ppt.LoadFromFile(inputFileName);

// Save as PDF
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.PDF });

// Clean up resources
ppt.Dispose();
```

---

# spire.presentation javascript conversion
## convert one slide to SVG format
```javascript
// Create PPT document
let ppt = wasmModule.Presentation.Create();

// Load PPT file from VFS
ppt.LoadFromFile(inputFileName);

// Convert the second slide to SVG
let svgByte = ppt.Slides.get_Item(1).SaveToSVG();
// Define the output file name
const outputFileName = "OneSlideToSVG.svg";
// Save as svg
svgByte.Save(outputFileName);

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Conversion
## Convert PowerPoint slide to SVG format
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Save the slide to SVG bytes
let bytes = slide.SaveToSVG();

// Save as svg
bytes.Save(outputFileName);
```

---

# Spire.Presentation JavaScript Conversion
## Convert specific slide to PDF
```javascript
// Create PPT document
let ppt = wasmModule.Presentation.Create();

// Load the PPT document from VFS
ppt.LoadFromFile(inputFileName);

// Get the second slide
let slide = ppt.Slides.get_Item(1);

// Define the output file name
const outputFileName = "SpecificSlideToPDF.pdf";

//Save to file
slide.SaveToFile(outputFileName, spirepresentation.FileFormat.PDF);

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Conversion
## Convert PowerPoint presentation to HTML format
```javascript
// Create an instance of presentation document
let ppt = wasmModule.Presentation.Create();

// Load file
ppt.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ToHTML.html";

// Save the document 
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# PowerPoint to Image Conversion
## Convert PowerPoint slides to PNG images
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

for (let i = 0; i < ppt.Slides.Count; i++) {
  let images = ppt.Slides.get_Item(i)._SaveAsImage1();
  let fileName = `ToImage_img_${i}.png`;

  // Save each image in virtual storage
  images.Save(fileName);
  
  images.Dispose();
}
// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Conversion
## Convert PowerPoint to Markdown
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load PPT file from disk
ppt.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ToMarkDown.md";

// Save the document 
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Markdown });

// Clean up resources
ppt.Dispose();
```

---

# spire.presentation javascript conversion
## convert PowerPoint presentation to PDF
```javascript
//Create a PPT document
let ppt =wasmModule.Presentation.Create();

//Load PPT file from disk
ppt.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ToPDF.pdf";

// Save the document 
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.PDF });

// Clean up resources
ppt.Dispose();
```

---

# spire.presentation javascript pdf conversion
## convert PPT to PDF with specific page size
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load PPT file from disk
ppt.LoadFromFile(inputFileName);

//Set A4 page size
ppt.SlideSize.Type = wasmModule.SlideSizeType.A4;

//Set landscape orientation
ppt.SlideSize.Orientation = wasmModule.SlideOrienation.Landscape;

// Define the output file name
const outputFileName = "ToPdfWithSpecificPageSize.pdf";

// Save the document 
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.PDF });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Conversion
## Convert PPT document to PPTX format
```javascript
// Create PPT document
let ppt = wasmModule.Presentation.Create();

// Load the PPT file from disk
ppt.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ToPPTX.pptx";

// Save the document 
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# PowerPoint to Image Conversion
## Convert PowerPoint slide to image with specific dimensions
```javascript
// Create an instance of presentation document
let ppt = wasmModule.Presentation.Create();

// Load file
ppt.LoadFromFile(inputFileName);

// Convert slide to image with specific size
let images = ppt.Slides.get_Item(0)._SaveAsImage(600, 400);

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Conversion
## Convert PowerPoint slides to SVG format
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

for (let i = 0; i < ppt.Slides.Count; i++) {
  let svgBytes = ppt.Slides.get_Item(i).SaveToSVG();
  let fileName = `ToSVG-${i}.svg`;

  // Save each image in virtual storage
  svgBytes.Save(fileName);
  
  svgBytes.Dispose();
}
// Clean up resources
ppt.Dispose();
```

---

# spire presentation javascript conversion
## convert PPT to XPS format
```javascript
// Create an instance of presentation document
let ppt = wasmModule.Presentation.Create();

// Load file
ppt.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ToXPS.xps";

// Save the document 
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.XPS });

// Clean up resources
ppt.Dispose();
```

---

# spire.presentation javascript image
## change image size in PowerPoint presentation
```javascript
// Change image size
let scale = 0.5;
for (let i = 0; i < ppt.Slides.Count; i++) {
  let slide = ppt.Slides.get_Item(i);
  for (let j = 0; j < slide.Shapes.Count; j++) {
    let shape = slide.Shapes.get_Item(j);
    if (shape instanceof wasmModule.SlidePicture) {
      let image = shape;
      image.Width = image.Width * scale;
      image.Height = image.Height * scale;
    }
  }
}
```

---

# spire presentation javascript image cropping
## crop image in powerpoint presentation
```javascript
// Get the first shape in first slide
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// If the shape is SlidePicture
if (shape instanceof wasmModule.SlidePicture) {
  let slidePicture = shape;
  //Crop image
  slidePicture.Crop(slidePicture.Left + 50, slidePicture.Top + 50, 100, 200);
}
```

---

# spire presentation javascript image extraction
## extract images from PowerPoint presentation
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

for (let i = 0; i < ppt.Images.Count; i++) {
  let image = ppt.Images.get_Item(i).Image;
  let imageName = `Images_${i}.png`;

  // Save each image in virtual storage
  image.Save(imageName);
  const imageFileArray = wasmModule.FS.readFile(imageName);
  const imageBlob = new Blob([imageFileArray], { type: "image/png" });

  // Add each image URL to the array for download
  imageDownloads.value.push({
    name: imageName,
    url: URL.createObjectURL(imageBlob),
  });

  image.Dispose();
}

// Clean up resources
ppt.Dispose();
```

---

# Extract Images from Specific PowerPoint Slide
## This code demonstrates how to extract images from a specific slide in a PowerPoint presentation

```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

let i = 0;
// Traverse all shapes in the second slide
for (let j = 0; j < ppt.Slides.get_Item(1).Shapes.Count; j++) {
  let shape = ppt.Slides.get_Item(1).Shapes.get_Item(j);
  // It is the SlidePicture object
  if (shape instanceof wasmModule.SlidePicture) {
    // Save to image
    let ps = shape;
    let fileName = `SlidePic_${i}.png`;

    let image = ps.PictureFill.Picture.EmbedImage.Image;
    image.Save(fileName);
    
    image.Dispose();
    i++;
  }
}

// Clean up resources
ppt.Dispose();
```

---

# spire.presentation javascript emf image
## insert EMF image in PowerPoint slide
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

// Define image size
let rect = wasmModule.RectangleF.FromLTRB(100, 100, 580, 460);

//Append the EMF in slide
let image = ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: ImageFileName, rectangle: rect });
image.Line.FillType = wasmModule.FillFormatType.None;
```

---

# Spire.Presentation JavaScript Image Insertion
## Insert an image into a PowerPoint presentation
```javascript
// Insert image to PPT
let rect1 = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 280, 140, (120 + ppt.SlideSize.Size.Width / 2 - 280), 260);
let image = ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: ImageFileName, rectangle: rect1 });
image.Line.FillType = wasmModule.FillFormatType.None;
```

---

# spire.presentation javascript image removal
## remove images from powerpoint slides
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

// Get the first slide
let slide = ppt.Slides.get_Item(0);

for (let i = slide.Shapes.Count - 1; i >= 0; i--) {
  //It is the SlidePicture object
  if (slide.Shapes.get_Item(i) instanceof wasmModule.SlidePicture) {
    slide.Shapes.RemoveAt(i);
  }
}

// Save the document 
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# spire.presentation javascript image frame format
## set format of image frame in powerpoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load image as stream
let stream = wasmModule.Stream.CreateByFile(imageFileName);
let imageData = ppt.Images.Append({ stream: stream });

// Add the image in document
let rect = wasmModule.RectangleF.FromLTRB(100, 100, (imageData.Width / 2 + 100), (imageData.Height / 2 + 100));
let pptImage = ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, embedImage: imageData, rectangle: rect });

// Set the formatting of the image frame
pptImage.Line.FillFormat.FillType = wasmModule.FillFormatType.Solid;
pptImage.Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_LightBlue();
pptImage.Line.Width = 5;
pptImage.Rotation = -45;
```

---

# spire presentation javascript image transparency
## set image transparency in powerpoint
```javascript
// Create an instance of presentation document
let ppt = wasmModule.Presentation.Create();

//Add a shape
let rect1 = wasmModule.RectangleF.FromLTRB(200, 100, 450, 350);
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({ shapeType: wasmModule.ShapeType.Rectangle, rectangle: rect1 });
shape.Line.FillType = wasmModule.FillFormatType.None;
//Fill shape with image
shape.Fill.FillType = wasmModule.FillFormatType.Picture;
shape.Fill.PictureFill.Picture.Url = imagePathName;
shape.Fill.PictureFill.FillType = wasmModule.PictureFillType.Stretch;
//Set transparency on image
shape.Fill.PictureFill.Picture.Transparency = 50;
```

---

# Spire.Presentation JavaScript Image Update
## Replace image in PowerPoint presentation
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Append a new image to replace an existing image
let stream = wasmModule.Stream.CreateByFile(fileStreamName);
let image = ppt.Images.Append({ stream: stream });
stream.Close();

// Replace the image which title is "image1" with the new image
for (let i = 0; i < slide.Shapes.Count; i++) {
  let shape = slide.Shapes.get_Item(i);
  if (shape instanceof wasmModule.SlidePicture) {
    if (shape.AlternativeTitle == "image1") {
      shape.PictureFill.Picture.EmbedImage = image;
    }
  }
}
```

---

# Adding Image to Table Cell in Presentation
## This code demonstrates how to add an image into a table cell in a PowerPoint presentation
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

// Get the first shape
let table = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Load the image and insert it into table cell
let stream = wasmModule.Stream.CreateByFile(fileStreamName);
let pptImg = ppt.Images.Append({ stream: stream });
stream.Close();

table.get_Item(1, 1).FillFormat.FillType = wasmModule.FillFormatType.Picture;
table.get_Item(1, 1).FillFormat.PictureFill.Picture.EmbedImage = pptImg;
table.get_Item(1, 1).FillFormat.PictureFill.FillType = wasmModule.PictureFillType.Stretch;
```

---

# Spire.Presentation JavaScript Table
## Add a row to a table in PowerPoint document
```javascript
// Create a PowerPoint document
let ppt = wasmModule.Presentation.Create();

// Load the file
ppt.LoadFromFile(inputFileName);

// Get the table within the PowerPoint document
let table = ppt.Slides.get_Item(0).Shapes.get_Item(0);

// Get the first row
let row = table.TableRows.get_Item(1);

// Clone the row and add it to the end of table
table.TableRows.Append(row);
let rowCount = table.TableRows.Count;

// Get the last row
let lastRow = table.TableRows.get_Item(rowCount - 1);

// Set new data of the first cell of last row
lastRow.get_Item(0).TextFrame.Text = " The first added cell";

// Set new data of the second cell of last row
lastRow.get_Item(1).TextFrame.Text = " The second added cell";

// Save the document
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Table
## Adjust column width by text width
```javascript
// Get the table from the first slide of the sample document.
let slide = ppt.Slides.get_Item(0);
let table = slide.Shapes.get_Item(0);

// Adjust the first column width of table by text width.
table.ColumnsList.get_Item(0).AdjustColumnByTextWidth();
```

---

# spire.presentation javascript table
## clone row and column in table
```javascript
// Create PPT document 
let ppt = wasmModule.Presentation.Create();

// Access first slide
let sld = ppt.Slides.get_Item(0);

// Define columns with widths and rows with heights
let widths = [110, 110, 110];
let heights = [50, 30, 30, 30, 30];

// Add table shape to slide
let table = ppt.Slides.get_Item(0).Shapes.AppendTable(ppt.SlideSize.Size.Width / 2 - 275, 90, widths, heights);

// Add text to the row 1 cell 1
table.get_Item(0, 0).TextFrame.Text = "Row 1 Cell 1";

// Add text to the row 1 cell 2
table.get_Item(1, 0).TextFrame.Text = "Row 1 Cell 2";

// Clone row 1 at end of table
table.TableRows.Append(table.TableRows.get_Item(0));

// Add text to the row 2 cell 1
table.get_Item(0, 1).TextFrame.Text = "Row 2 Cell 1";

// Add text to the row 2 cell 2
table.get_Item(1, 1).TextFrame.Text = "Row 2 Cell 2";

// Clone row 2 as the 4th row of table
table.TableRows.Insert(3, table.TableRows.get_Item(1));

// Clone column 1 at end of table
table.ColumnsList.Add(table.ColumnsList.get_Item(0));

// Clone the 2nd column at 4th column index
table.ColumnsList.Insert(3, table.ColumnsList.get_Item(1));
```

---

# PowerPoint Table Creation
## Create and format a table in a PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Define table dimensions
let widths = [100, 100, 150, 100, 100];
let heights = [15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15];

// Add new table to PPT
let table = ppt.Slides.get_Item(0).Shapes.AppendTable(ppt.SlideSize.Size.Width / 2 - 275, 90, widths, heights);
let dataStr = [
  ["Name", "Capital", "Continent", "Area", "Population"],
  ["Venezuela", "Caracas", "South America", "912047", "19700000"],
  ["Bolivia", "La Paz", "South America", "1098575", "7300000"],
  ["Brazil", "Brasilia", "South America", "8511196", "150400000"],
  ["Canada", "Ottawa", "North America", "9976147", "26500000"],
  ["Chile", "Santiago", "South America", "756943", "13200000"],
  ["Colombia", "Bagota", "South America", "1138907", "33000000"],
  ["Cuba", "Havana", "North America", "114524", "10600000"],
  ["Ecuador", "Quito", "South America", "455502", "10600000"],
  ["Paraguay", "Asuncion", "South America", "406576", "4660000"],
  ["Peru", "Lima", "South America", "1285215", "21600000"],
  ["Jamaica", "Kingston", "North America", "11424", "2500000"],
  ["Mexico", "Mexico City", "North America", "1967180", "88600000"]
];

// Add data to table
for (let i = 0; i < 13; i++) {
  for (let j = 0; j < 5; j++) {
    // Fill the table with data
    table.get_Item(j, i).TextFrame.Text = dataStr[i][j];
    
    // Set the Font
    table.get_Item(j, i).TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Arial Narrow");
  }
}

// Set the alignment of the first row to Center
for (let i = 0; i < 5; i++) {
  table.get_Item(i, 0).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Center;
}

// Set the style of table
table.StylePreset = wasmModule.TableStylePreset.LightStyle3Accent1;
```

---

# PowerPoint Table Data and Style Editor
## Edit table data and style in PowerPoint document
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

// Store the data used in replacement in string array
let str = ["Germany", "Berlin", "Europe", "0152458", "20860000"];

// Get the table in PowerPoint document
for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
  let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
  if (shape instanceof wasmModule.ITable) {
    // Change the style of table
    shape.StylePreset = wasmModule.TableStylePreset.LightStyle1Accent2;

    for (let i = 0; i < shape.ColumnsList.Count; i++) {
      // Replace the data in cell
      shape.get_Item(i, 2).TextFrame.Text = str[i];

      // Set the highlight color
      shape.get_Item(i, 2).TextFrame.TextRange.HighlightColor.Color = wasmModule.Color.get_BlueViolet();
    }
  }
}

// Save the document 
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Table
## Fill all table cells with color
```javascript
// Fill the table cells with color.
for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
  let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
  if (shape instanceof wasmModule.ITable) {
    let table = shape;
    for (let j = 0; j < table.TableRows.Count; j++) {
      let row = table.TableRows.get_Item(j);
      for (let k = 0; k < row.Count; k++) {
        let cell = row.get_Item(k);
        cell.FillFormat.FillType = wasmModule.FillFormatType.Solid;
        cell.FillFormat.SolidColor.Color = wasmModule.Color.get_Pink();
      }
    }
  }
}
```

---

# Spire.Presentation JavaScript Table
## Fill particular table row with color
```javascript
// Fill particular table row with color.
for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
  let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
  if (shape instanceof wasmModule.ITable) {
    let table = shape;
    let row = table.TableRows.get_Item(1);
    for (let k = 0; k < row.Count; k++) {
      let cell = row.get_Item(k);
      cell.FillFormat.FillType = wasmModule.FillFormatType.Solid;
      cell.FillFormat.SolidColor.Color = wasmModule.Color.get_Pink();
    }
  }
}
```

---

# PowerPoint Table Cell Border Color Extraction
## Get display color and border color of table cells in PowerPoint document
```javascript
// Create PPT document and load file
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

//Get the table in the first slide
let table = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Get borders' color of the first cell
let sb = [];
sb.push("Color of left border:" + table.get_Item(0, 0).BorderLeftDisplayColor.ToString());
sb.push("Color of top border:" + table.get_Item(0, 0).BorderTopDisplayColor.ToString());
sb.push("Color of right border:" + table.get_Item(0, 0).BorderRightDisplayColor.ToString());
sb.push("Color of bottom border:" + table.get_Item(0, 0).BorderBottomDisplayColor.ToString());

//Get display color of the first cell
sb.push("Color of cell:" + table.get_Item(0, 0).DisplayColor.ToString());

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Table
## Identify merged cells in a PowerPoint table
```javascript
// Get the first slide
let slide = ppt.Slides.get_Item(0);
let str = [];
let output = "";
for (let i = 0; i < slide.Shapes.Count; i++) {
  let shape = slide.Shapes.get_Item(i);
  // Verify if it is table
  if (shape instanceof wasmModule.ITable) {
    let table = shape;
    for (let r = 0; r < table.TableRows.Count; r++) {
      for (let c = 0; c < table.ColumnsList.Count; c++) {
        // Get cell
        let currentCell = table.TableRows.get_Item(r).get_Item(c);
        // Identify if it is merged cell
        if (currentCell.RowSpan > 1 || currentCell.ColSpan > 1) {
          output = `Cell ${r}:${c} is a part of merged cell with RowSpan=${currentCell.RowSpan} and ColSpan=${currentCell.ColSpan} starting from Cell ${currentCell.FirstRowIndex}:${currentCell.FirstColumnIndex}.`;
          str.push(output);
        }
      }
    }
  }
}
```

---

# spire presentation javascript table
## lock aspect ratio of table
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load PPT file from disk
ppt.LoadFromFile(inputFileName);
//Get the first slide
let slide = ppt.Slides.get_Item(0);

for (let i = 0; i < slide.Shapes.Count; i++) {
    let shape = slide.Shapes.get_Item(i);
    //Verify if it is table
    if (shape instanceof wasmModule.ITable) {
        let table = shape;
        //Lock aspect ratio
        table.ShapeLocking.AspectRatioProtection = true;
    }
}
```

---

# spire.presentation javascript table
## merge table cells in powerpoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

ppt.LoadFromFile(inputFileName);

for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
    let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
    // Verify if it is table
    if (shape instanceof wasmModule.ITable) {
        let table = shape;
        // Merge the second row and third row of the first column
        table.MergeCells(table.get_Item(0,1), table.get_Item(0,2), false);

        table.MergeCells(table.get_Item(3,4), table.get_Item(4,4), true);
    }
}

// Save the document to the specified path
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Table Manipulation
## Remove rows and columns from a table in a PowerPoint presentation
```javascript
//Get the table in PPT document
for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
    let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
    //Verify if it is table
    if (shape instanceof wasmModule.ITable) {
        let table = shape;
        //Remove the second column
        table.ColumnsList.RemoveAt(1, false);

        //Remove the second row
        table.TableRows.RemoveAt(1, false);
    }
}
```

---

# spire.presentation javascript table
## remove table border style in powerpoint
```javascript
// Iterate through all slides
for (let o = 0; o < ppt.Slides.Count; o++) {
    let slide = ppt.Slides.get_Item(o);
    // Iterate through all shapes in the slide
    for (let i = 0; i < slide.Shapes.Count; i++) {
        let shape = slide.Shapes.get_Item(i);
        // Verify if it is table
        if (shape instanceof wasmModule.ITable) {
            let table = shape;
            // Iterate through all rows in the table
            for (let j = 0; j < table.TableRows.Count; j++) {
                let row = table.TableRows.get_Item(j);
                // Iterate through all cells in the row
                for (let k = 0; k < row.Count; k++) {
                    let cell = row.get_Item(k);
                    // Remove border styles for each cell
                    cell.BorderTop.FillType = wasmModule.FillFormatType.None;
                    cell.BorderBottom.FillType = wasmModule.FillFormatType.None;
                    cell.BorderLeft.FillType = wasmModule.FillFormatType.None;
                    cell.BorderRight.FillType = wasmModule.FillFormatType.None;
                }
            }
        }
    }
}
```

---

# spire.presentation javascript table
## remove tables from powerpoint slide
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the file from disk.
ppt.LoadFromFile(inputFileName);

//Get the tables within the PPT document.
let shape_tems = [];

for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
    let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
    if (shape instanceof wasmModule.ITable) {
        //Add new table to table list.
        shape_tems.push(shape);
    }
}

//Remove all the tables form the first slide.
for (let i = 0; i < shape_tems.length; i++) {
    let shape = shape_tems[i];
    ppt.Slides.get_Item(0).Shapes.Remove(shape);
}
```

---

# spire.presentation javascript table alignment
## set horizontal and vertical alignment for table cells
```javascript
for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
    let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
    if (shape instanceof wasmModule.ITable) {
        let table = shape;
        //Horizontal Alignment
        //Set the horizontal alignment for the cells in first column
        table.get_Item(0,1).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Left;
        table.get_Item(0,2).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Center;
        table.get_Item(0,3).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Right;
        table.get_Item(0,4).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Justify;

        //Vertical Alignment
        //Set the vertical alignment for the cells in second column
        table.get_Item(1,1).TextAnchorType = wasmModule.TextAnchorType.Top;
        table.get_Item(1,2).TextAnchorType = wasmModule.TextAnchorType.Center;
        table.get_Item(1,3).TextAnchorType = wasmModule.TextAnchorType.Bottom;
        table.get_Item(1,4).TextAnchorType = wasmModule.TextAnchorType.None;

        //Both orientations
        //Set the both horizontal and vertical alignment for the cells in the third column
        table.get_Item(2,1).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Left;
        table.get_Item(2,1).TextAnchorType = wasmModule.TextAnchorType.Top;

        table.get_Item(2,2).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Right;
        table.get_Item(2,2).TextAnchorType = wasmModule.TextAnchorType.Center;

        table.get_Item(2,3).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Justify;
        table.get_Item(2,3).TextAnchorType = wasmModule.TextAnchorType.Bottom;

        table.get_Item(2,4).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Center;
        table.get_Item(2,4).TextAnchorType = wasmModule.TextAnchorType.Top;
    }
}
```

---

# spire presentation javascript table borders
## set border type and color for existing table
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the file from disk.
ppt.LoadFromFile(inputFileName);

//Get the table from the first slide of the sample document.
let slide = ppt.Slides.get_Item(0);
let table = slide.Shapes.get_Item(0);

//Set the border type as Inside and the border color as blue.
table.SetTableBorder(wasmModule.TableBorderType.Inside, 1, wasmModule.Color.get_Blue());
```

---

# spire presentation javascript table borders
## set border type and color for newly added tables
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Set the table width and height for each table cell
let tableWidth = [100, 100, 100, 100, 100]
let tableHeight = [20, 20];

// Traverse all the border type of the table
for (let item in wasmModule.TableBorderType) {
    // Add a table to the presentation slide with the setting width and height
    let itable = ppt.Slides.Append().Shapes.AppendTable(100, 100, tableWidth, tableHeight);

    // Add some text to the table cell
    itable.TableRows.get_Item(0).get_Item(0).TextFrame.Text = "Row";
    itable.TableRows.get_Item(1).get_Item(0).TextFrame.Text = "Column";

    // Set the border type, border width and the border color for the table
    itable.SetTableBorder(wasmModule.TableBorderType[item], 1.5, wasmModule.Color.get_Red());
}
```

---

# Spire.Presentation JavaScript Table
## Set first row as header in PowerPoint table
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

ppt.LoadFromFile(inputFileName);

let table = null;
for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
    let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
    if (shape instanceof wasmModule.ITable) {
        table = shape;
    }
}

table.FirstRow = true;
```

---

# Spire.Presentation JavaScript Table
## Set Row Height and Column Width for Table
```javascript
//Get the table
let table = null;
for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
    let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
    if (shape instanceof wasmModule.ITable) {
        table = shape;
        //Set the height for the rows
        table.TableRows.get_Item(0).Height = 100;
        table.TableRows.get_Item(1).Height = 80;
        table.TableRows.get_Item(2).Height = 60;
        table.TableRows.get_Item(3).Height = 40;
        table.TableRows.get_Item(4).Height = 20;

        //Set the column width
        table.ColumnsList.get_Item(0).Width = 60;
        table.ColumnsList.get_Item(1).Width = 80;
        table.ColumnsList.get_Item(2).Width = 120;
        table.ColumnsList.get_Item(3).Width = 140;
        table.ColumnsList.get_Item(4).Width = 160;
    }
}
```

---

# PowerPoint Table Border Style
## Set solid border style for all cells in tables within a PowerPoint presentation
```javascript
// Find the table by looping through all the slides, and then set borders for it
for (let o = 0; o < ppt.Slides.Count; o++) {
    let slide = ppt.Slides.get_Item(o);
    for (let i = 0; i < slide.Shapes.Count; i++) {
        let shape = slide.Shapes.get_Item(i);
        // Verify if it is table
        if (shape instanceof wasmModule.ITable) {
            let table = shape;
            for (let j = 0; j < table.TableRows.Count; j++) {
                let row = table.TableRows.get_Item(j);
                for (let k = 0; k < row.Count; k++) {
                    let cell = row.get_Item(k);
                    cell.BorderTop.FillType = wasmModule.FillFormatType.Solid;
                    cell.BorderBottom.FillType = wasmModule.FillFormatType.Solid;
                    cell.BorderLeft.FillType = wasmModule.FillFormatType.Solid;
                    cell.BorderRight.FillType = wasmModule.FillFormatType.Solid;
                }
            }
        }
    }
}
```

---

# Spire.Presentation JavaScript Table Style
## Apply built-in style to table in PowerPoint presentation
```javascript
//Get the table
let table = null;
for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
    let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
    if (shape instanceof wasmModule.ITable) {
        table = shape;
        //Set the table style from TableStylePreset and apply it to selected table
        table.StylePreset = wasmModule.TableStylePreset.MediumStyle1Accent2;
    }
}
```

---

# spire presentation javascript table text formatting
## set text format for table cells in a presentation
```javascript
//Get the first slide
let slide = ppt.Slides.get_Item(0);

for (let i = 0; i < slide.Shapes.Count; i++) {
    let shape = slide.Shapes.get_Item(i);
    //Verify if it is table
    if (shape instanceof wasmModule.ITable){
        let table = shape;
        let cell1 = table.TableRows.get_Item(0).get_Item(0);
        //Set table cell's text alignment type
        cell1.TextAnchorType = wasmModule.TextAnchorType.Top;
        //Set italic style
        cell1.TextFrame.TextRange.Format.IsItalic = wasmModule.TriState.True;

        let cell2 = table.TableRows.get_Item(1).get_Item(0);
        //Set table cell's foreground color
        cell2.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
        cell2.TextFrame.TextRange.Fill.SolidColor.Color = wasmModule.Color.get_Green();
        //Set table cell's background color
        cell2.FillFormat.FillType = wasmModule.FillFormatType.Solid;
        cell2.FillFormat.SolidColor.Color = wasmModule.Color.get_LightGray();

        let cell3 = table.TableRows.get_Item(2).get_Item(2);
        //Set table cell's font and font size
        cell3.TextFrame.TextRange.FontHeight = 12;
        cell3.TextFrame.TextRange.LatinFont = wasmModule.TextFont.Create("Arial Black");
        cell3.TextFrame.TextRange.HighlightColor.Color = wasmModule.Color.get_YellowGreen();

        let cell4 = table.TableRows.get_Item(2).get_Item(1);
        //Set table cell's margin and borders
        cell4.MarginLeft = 20;
        cell4.MarginTop = 30;
        cell4.BorderTop.FillType = wasmModule.FillFormatType.Solid;
        cell4.BorderTop.SolidFillColor.Color = wasmModule.Color.get_Red();
        cell4.BorderBottom.FillType = wasmModule.FillFormatType.Solid;
        cell4.BorderBottom.SolidFillColor.Color = wasmModule.Color.get_Red();
        cell4.BorderLeft.FillType = wasmModule.FillFormatType.Solid;
        cell4.BorderLeft.SolidFillColor.Color = wasmModule.Color.get_Red();
        cell4.BorderRight.FillType = wasmModule.FillFormatType.Solid;
        cell4.BorderRight.SolidFillColor.Color = wasmModule.Color.get_Red();
    }
}
```

---

# PowerPoint Table Cell Splitting
## Split a specific table cell in PowerPoint presentation
```javascript
//Get the first slide.
let slide = ppt.Slides.get_Item(0);

//Get the table.
let table = slide.Shapes.get_Item(0);

//Split cell [1, 2] into 3 rows and 2 columns.
table.get_Item(1,2).Split(3, 2);
```

---

# Spire.Presentation JavaScript Table
## Traverse through cells of a table in PowerPoint
```javascript
//Get the table.
let table = null;
for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
    let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
    if (shape instanceof wasmModule.ITable) {
        table = shape;
        //Traverse through the cells of table.
        for (let j = 0; j < table.TableRows.Count; j++) {
            let row = table.TableRows.get_Item(j);
            for (let k = 0; k < row.Count; k++) {
                let cell = row.get_Item(k);
                // Get text from cell
                let cellText = cell.TextFrame.Text;
            }
        }
    }
}
```

---

# spire.presentation javascript hyperlink
## add hyperlink to image in PowerPoint
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Get the first slide
let slide = ppt.Slides.get_Item(0);

let rect = wasmModule.RectangleF.FromLTRB(480, 350, 640, 510);
let image = slide.Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName: imageFileName,rectangle: rect});

// Add hyperlink to the image
let hyperlink = wasmModule.ClickHyperlink.Create("https://www.e-iceblue.com");
image.Click = hyperlink;
```

---

# spire.presentation javascript hyperlink
## add hyperlink to SmartArt nodes in PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

ppt.LoadFromFile(inputFileName);

//Get the smartArt shape
let sr = ppt.Slides.get_Item(0).Shapes.get_Item(0);
//Add hyperlinks to the nodes
let node = sr.Nodes.get_Item(0);
node.Click = wasmModule.ClickHyperlink.Create_silde(ppt.Slides.get_Item(1));
node = sr.Nodes.get_Item(1);
node.Click = wasmModule.ClickHyperlink.Create_silde(ppt.Slides.get_Item(2));
node = sr.Nodes.get_Item(2);
node.Click = wasmModule.ClickHyperlink.Create_silde(ppt.Slides.get_Item(3));
```

---

# Spire.Presentation JavaScript Hyperlink
## Add hyperlink to text in PowerPoint
```javascript
//Find the text we want to add link to it.
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(0);
let tp = shape.TextFrame.TextRange.Paragraph;
let temp = tp.Text;

//Split the original text.
let textToLink = "Spire.Presentation";
let strSplit = temp.split(textToLink);

//Clear all text.
tp.TextRanges.Clear();

//Add new text.
let tr = wasmModule.TextRange.Create(strSplit[0]);
tp.TextRanges.Append(tr);

//Add the hyperlink.
tr = wasmModule.TextRange.Create(textToLink);
tr.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html";
tp.TextRanges.Append(tr);
```

---

# Spire.Presentation JavaScript Hyperlink
## Change hyperlink color in PowerPoint
```javascript
// Get the first slide
let slide = ppt.Slides.get_Item(0);

// Get the theme of the slide
let theme = slide.Theme;

// Change the color of hyperlink to red
theme.ColorScheme.HyperlinkColor.Color = wasmModule.Color.get_Red();
```

---

# Get Linked Slide from PowerPoint Shape
## Retrieve the target slide information from a shape with hyperlink action
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load ppt file
ppt.LoadFromFile(inputFileName);

// Get the second slide
let slide = ppt.Slides.get_Item(1);

// Get the first shape of the second slide
let shape = slide.Shapes.get_Item(0);

// Get the linked slide index
if (shape.Click.ActionType == wasmModule.HyperlinkActionType.GotoSlide) {
    let targetSlide = shape.Click.TargetSlide;
    strB.push("Linked slide number = " + targetSlide.SlideNumber);
}
```

---

# Spire.Presentation JavaScript Hyperlink
## Add hyperlink and set outline style in PowerPoint presentation
```javascript
//Add new shape to PPT document
let rec = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 255, 120, (400 + ppt.SlideSize.Size.Width / 2 - 255), 220);
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:rec});
shape.Fill.FillType = wasmModule.FillFormatType.None;
shape.Line.FillType = wasmModule.FillFormatType.None;

//Add a paragraph with hyperlink
let para1 = wasmModule.TextParagraph.Create();
let tr1 = wasmModule.TextRange.Create("Click to know more about Spire.Presentation");
tr1.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html";
para1.TextRanges.Append(tr1);

//Set the format of textrange
tr1.Format.FontHeight = 20;
tr1.IsItalic = wasmModule.TriState.True;

//Set the outline format of textrange
tr1.TextLineFormat.FillFormat.FillType = wasmModule.FillFormatType.Solid;
tr1.TextLineFormat.FillFormat.SolidFillColor.Color = wasmModule.Color.get_LightSeaGreen();
tr1.TextLineFormat.JoinStyle = wasmModule.LineJoinType.Round;
tr1.TextLineFormat.Width = 2;

//Add the paragraph to shape
shape.TextFrame.Paragraphs._Append(para1);
shape.TextFrame.Paragraphs._Append(wasmModule.TextParagraph.Create());
```

---

# spire.presentation javascript hyperlinks
## create hyperlinks in a PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Add new shape to PPT document
let rec = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 255, 120, (500 + ppt.SlideSize.Size.Width / 2 - 255), 400);
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: rec});
shape.Fill.FillType = wasmModule.FillFormatType.None;
shape.Line.Width = 0;

//Add some paragraphs with hyperlinks
let para1 = wasmModule.TextParagraph.Create();
let tr = wasmModule.TextRange.Create("E-iceblue");
tr.Fill.FillType = wasmModule.FillFormatType.Solid;
tr.Fill.SolidColor.Color = wasmModule.Color.get_Blue();
para1.TextRanges.Append(tr);
para1.Alignment = wasmModule.TextAlignmentType.Center;
shape.TextFrame.Paragraphs._Append(para1);
shape.TextFrame.Paragraphs._Append(wasmModule.TextParagraph.Create());

//Add hyperlink paragraph
let para2 = wasmModule.TextParagraph.Create();
let tr1 = wasmModule.TextRange.Create("Click to know more about Spire.Presentation.");
tr1.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html";
para2.TextRanges.Append(tr1);
shape.TextFrame.Paragraphs._Append(para2);
shape.TextFrame.Paragraphs._Append(wasmModule.TextParagraph.Create());

//Add hyperlink paragraph
let para3 = wasmModule.TextParagraph.Create();
let tr2 = wasmModule.TextRange.Create("Click to visit E-iceblue Home page.");
tr2.ClickAction.Address = "https://www.e-iceblue.com/";
para3.TextRanges.Append(tr2);
shape.TextFrame.Paragraphs._Append(para3);
shape.TextFrame.Paragraphs._Append(wasmModule.TextParagraph.Create());

//Add hyperlink paragraph
let para4 = wasmModule.TextParagraph.Create();
let tr3 = wasmModule.TextRange.Create("Click to go to the forum to raise questions.");
tr3.ClickAction.Address = "https://www.e-iceblue.com/forum/components-f5.html";
para4.TextRanges.Append(tr3);
shape.TextFrame.Paragraphs._Append(para4);
shape.TextFrame.Paragraphs._Append(wasmModule.TextParagraph.Create());

//Add email hyperlink paragraph
let para5 = wasmModule.TextParagraph.Create();
let tr4 = wasmModule.TextRange.Create("Click to contact our sales team via email.");
tr4.ClickAction.Address = "mailto:sales@e-iceblue.com";
para5.TextRanges.Append(tr4);
shape.TextFrame.Paragraphs._Append(para5);
shape.TextFrame.Paragraphs._Append(wasmModule.TextParagraph.Create());

//Add email hyperlink paragraph
let para6 = wasmModule.TextParagraph.Create();
let tr5 = wasmModule.TextRange.Create("Click to contact our support team via email.");
tr5.ClickAction.Address = "mailto:support@e-iceblue.com";
para6.TextRanges.Append(tr5);
shape.TextFrame.Paragraphs._Append(para6);

for (let i = 0; i < shape.TextFrame.Paragraphs.Count; i++) {
    let para = shape.TextFrame.Paragraphs.get_Item(i);
    if(para.Text.length != 0){
        para.TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");
        para.TextRanges.get_Item(0).FontHeight = 20;
    }
}
```

---

# PowerPoint Hyperlink to Specific Slide
## Create a hyperlink that links to a specific slide in a PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Append a slide to it.
ppt.Slides.Append();

//Add a shape to the second slide.
let shape = ppt.Slides.get_Item(1).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(10, 50, 210, 100)});
shape.Fill.FillType = wasmModule.FillFormatType.None;
shape.Line.FillType = wasmModule.FillFormatType.None;
shape.TextFrame.Text = "Jump to the first slide";

//Create a hyperlink based on the shape and the text on it, linking to the first slide.
let hyperlink = wasmModule.ClickHyperlink.Create_silde(ppt.Slides.get_Item(0));
shape.Click = hyperlink;
shape.TextFrame.TextRange.ClickAction = hyperlink;
```

---

# spire.presentation javascript hyperlink
## create hyperlink to last viewed slide
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

ppt.LoadFromFile(inputFileName);

let slide = ppt.Slides.get_Item(0);
// Draw a shape
let autoShape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: wasmModule.RectangleF.FromLTRB(100, 100, 200, 200)});
// Link to last viewed slide show
autoShape.Click = wasmModule.ClickHyperlink.get_LastVievedSlide();

// Save the document to the specified path
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# PowerPoint Hyperlink Modification
## Modify hyperlink address and text in a PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the file from disk.
ppt.LoadFromFile(inputFileName);

//Find the hyperlinks you want to edit.
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Edit the link text and the target URL.
shape.TextFrame.TextRange.ClickAction.Address = "http://www.e-iceblue.com";
shape.TextFrame.TextRange.Text = "E-iceblue";
```

---

# Spire.Presentation JavaScript Hyperlink
## Remove hyperlink from PowerPoint slide shape
```javascript
//Get the shape and its text with hyperlink.
let shape = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Set the ClickAction property into null to remove the hyperlink.
shape.TextFrame.TextRange.ClickAction = "";
```

---

# Extract Audio from Presentation
## Core functionality for extracting audio from PowerPoint presentations
```javascript
let inputFileName = "audio.pptx";
let outFileName = "ExtractAudio.wav";

// Create a PPT document
let ppt = wasmModule.Presentation.Create();
ppt.LoadFromFile(inputFileName);

// Iterate through shapes to find and extract audio
for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
    let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
    if(shape instanceof wasmModule.IAudio){
        let audio = shape;
        let AudioData = audio.Data;
        AudioData.SaveToFile(outFileName);        
    }
}

// Clean up resources
ppt.Dispose();
```

---

# Extract Videos from PowerPoint Presentation
## This code demonstrates how to extract embedded videos from a PowerPoint presentation by iterating through slides and shapes.
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the PPT document from disk.
ppt.LoadFromFile(inputFileName);

//Define a variable
let i = 0;

//String for output file
let result = `ExtractVideo_${i}.avi`;
//Traverse all the slides of PPT file
for (let j = 0; j < ppt.Slides.Count; j++) {
    let slide = ppt.Slides.get_Item(j);
    //Traverse all the shapes of slides
    for (let k = 0; k < slide.Shapes.Count; k++) {
        let shape = slide.Shapes.get_Item(k);
        //If shape is IVideo
        if (shape instanceof wasmModule.IVideo)
        {
            //Save the video
            shape.EmbeddedVideoData.SaveToFile(result);
            i++;
        }
    }
}
```

---

# Spire.Presentation JavaScript Audio
## Hide audio during presentation show
```javascript
//Get the first slide
let slide = ppt.Slides.get_Item(0);

//Hide Audio during show
for (let i = 0; i < slide.Shapes.Count; i++) {
    let shape = slide.Shapes.get_Item(i);
    if(shape instanceof wasmModule.IAudio){
        let audio = shape;
        audio.HideAtShowing = true;
    }
}
```

---

# Spire.Presentation JavaScript Audio
## Insert audio into PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load the document from disk
ppt.LoadFromFile(inputFileName);

// Add title
let rec_title = wasmModule.RectangleF.FromLTRB(50, 240, 210, 290);
let shape_title = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: rec_title});
shape_title.ShapeStyle.LineColor.Color = wasmModule.Color.get_Transparent();

shape_title.Fill.FillType = wasmModule.FillFormatType.None;
let para_title = wasmModule.TextParagraph.Create();
para_title.Text = "Audio:";
para_title.Alignment = wasmModule.TextAlignmentType.Center;
para_title.TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Myriad Pro Light");
para_title.TextRanges.get_Item(0).FontHeight = 32;
para_title.TextRanges.get_Item(0).IsBold = wasmModule.TriState.True;
para_title.TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
para_title.TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.FromArgb(68, 68, 68);
shape_title.TextFrame.Paragraphs._Append(para_title);

// Insert audio into the document
let audioRect = wasmModule.RectangleF.FromLTRB(220, 240, 300, 320);
ppt.Slides.get_Item(0).Shapes.AppendAudioMedia({filePath:audioFileName, rectangle:audioRect});
```

---

# Spire.Presentation JavaScript Video Insertion
## Insert video into PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the document from disk
ppt.LoadFromFile(inputFileName);

//Add title
let rec_title = wasmModule.RectangleF.FromLTRB(50, 280, 210, 330);
let shape_title = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: rec_title});
shape_title.ShapeStyle.LineColor.Color = wasmModule.Color.get_Transparent();

shape_title.Fill.FillType = wasmModule.FillFormatType.None;
let para_title = wasmModule.TextParagraph.Create();
para_title.Text = "Video:";
para_title.Alignment = wasmModule.TextAlignmentType.Center;
para_title.TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Myriad Pro Light");
para_title.TextRanges.get_Item(0).FontHeight = 32;
para_title.TextRanges.get_Item(0).IsBold = wasmModule.TriState.True;
para_title.TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
para_title.TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.FromArgb(68, 68, 68);
shape_title.TextFrame.Paragraphs._Append(para_title);

//Insert video into the document
let videoRect = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 125, 240, (150 + ppt.SlideSize.Size.Width / 2 - 125), 390);

let video = ppt.Slides.get_Item(0).Shapes.AppendVideoMedia({filePath:videoFileName, rectangle:videoRect});

video.PictureFill.Picture.Url = imageFileName;
```

---

# spire.presentation javascript audio
## obtain sound effect properties from animation in powerpoint
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load file
ppt.LoadFromFile(inputFileName);

//Get the first slide
let slide = ppt.Slides.get_Item(0);

//Get the audio in a time node
let audio = slide.Timeline.MainSequence.get_Item(0).TimeNodeAudios[0];

//Get the properties of the audio, such as sound name, volume or detect if it's mute
let sb = [];
sb.push("SoundName: " + audio.SoundName);
sb.push("Volume: " + audio.Volume);
sb.push("IsMute: " + audio.IsMute);
```

---

# Spire.Presentation JavaScript Video Replacement
## Replace existing video in PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the PPT document from disk.
ppt.LoadFromFile(inputFileName);

let videos = ppt.Videos;

//Traverse all the slides of PPT file
for (let i = 0; i < ppt.Slides.Count; i++) {
    let slide = ppt.Slides.get_Item(i);
    //Traverse all the shapes of slides
    for (let j = 0; j < slide.Shapes.Count; j++) {
        let shape = slide.Shapes.get_Item(j);
        //If shape is IVideo
        if (shape instanceof wasmModule.IVideo)
        {
            //Replace the video
            let video = shape;
            //Load the video document from disk.
            let bts = wasmModule.Stream.CreateByFile(inputFileName2);
            let videoData = videos.Append({stream:bts});
            video.EmbeddedVideoData = videoData;
        }
    }
}

// Save the document to the specified path
ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript Video
## Set play mode for video in PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load the file from disk.
ppt.LoadFromFile("Template_Ppt_8.pptx");

// Find the video by looping through all the slides and set its play mode as auto.
for (let i = 0; i < ppt.Slides.Count; i++) {
    let slide = ppt.Slides.get_Item(i);
    for (let j = 0; j < slide.Shapes.Count; j++) {
        let shape = slide.Shapes.get_Item(j);
        if (shape instanceof wasmModule.IVideo){
            shape.PlayMode = wasmModule.VideoPlayMode.Auto;
        }
    }
}

// Clean up resources
ppt.Dispose();
```

---

# Speaker Notes Management in PowerPoint
## Add and retrieve speaker notes from PowerPoint slides
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the file from disk.
ppt.LoadFromFile(inputFileName);

//Get the first slide and in the PowerPoint document.
let slide = ppt.Slides.get_Item(0);

//Get the NotesSlide in the first slide,if there is no notes, we need to add it firstly.
let ns = slide.NotesSlide;

if (ns.H == undefined) {
    ns = slide.AddNotesSlide();
}

//Add the text string as the notes.
ns.NotesTextFrame.Text = "Speak notes added by Spire.Presentation";

let content = [];
content.push("The speaker notes added by Spire.Presentation is: " + ns.NotesTextFrame.Text);

//Get the speaker notes and save to txt file.
FS.writeFile(outputDirectoryName+outputFile_txt, content.join(""));

// Define the output file name
const outputFileName = "AddAndGetSpeakerNotes_out.pptx";

// Save the document to the specified path
ppt.SaveToFile({ file: outputDirectoryName+outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

// Clean up resources
ppt.Dispose();
```

---

# PowerPoint Comment Addition
## Add a comment to a PowerPoint slide
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

ppt.LoadFromFile(inputFileName);

// Comment author
let author = ppt.CommentAuthors.AddAuthor("E-iceblue", "comment:");

// Add comment
ppt.Slides.get_Item(0).AddComment({
  author: author,
  text: "Add comment", 
  position: wasmModule.PointF.Create(18, 25),
  dateTime: wasmModule.DateTime.get_Now()
});
```

---

# spire.presentation javascript notes
## add notes to PowerPoint slides
```javascript
//Add note slide
let notesSlide = slide.AddNotesSlide();

//Add paragraph in the notesSlide
let paragraph = wasmModule.TextParagraph.Create();
paragraph.Text = "Tips for making effective presentations:";
notesSlide.NotesTextFrame.Paragraphs._Append(paragraph);

paragraph = wasmModule.TextParagraph.Create();
paragraph.Text = "Use the slide master feature to create a consistent and simple design template.";
notesSlide.NotesTextFrame.Paragraphs._Append(paragraph);
//Set the bullet type for the paragraph in notesSlide
notesSlide.NotesTextFrame.Paragraphs.get_Item(1).BulletType = wasmModule.TextBulletType.Numbered;
notesSlide.NotesTextFrame.Paragraphs.get_Item(1).BulletStyle = wasmModule.NumberedBulletStyle.BulletArabicPeriod;

paragraph = wasmModule.TextParagraph.Create();
paragraph.Text = "Simplify and limit the number of words on each screen.";
notesSlide.NotesTextFrame.Paragraphs._Append(paragraph);
notesSlide.NotesTextFrame.Paragraphs.get_Item(2).BulletType = wasmModule.TextBulletType.Numbered;
notesSlide.NotesTextFrame.Paragraphs.get_Item(2).BulletStyle = wasmModule.NumberedBulletStyle.BulletArabicPeriod;

paragraph = wasmModule.TextParagraph.Create();
paragraph.Text = "Use contrasting colors for text and background.";
notesSlide.NotesTextFrame.Paragraphs._Append(paragraph);
notesSlide.NotesTextFrame.Paragraphs.get_Item(3).BulletType = wasmModule.TextBulletType.Numbered;
notesSlide.NotesTextFrame.Paragraphs.get_Item(3).BulletStyle = wasmModule.NumberedBulletStyle.BulletArabicPeriod;
```

---

# spire.presentation javascript comment
## delete and replace comments in presentation
```javascript
//Replace the text in the comment
ppt.Slides.get_Item(0).Comments[1].Text = "Replace comment";

//Delete the third comment
ppt.Slides.get_Item(0).DeleteComment({comment:ppt.Slides.get_Item(0).Comments[2]});
```

---

# Extract Comments from Presentation
## Extract text comments from PowerPoint presentation slides
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the file from disk.
ppt.LoadFromFile(inputFileName);

let str = [];

//Get all comments from the first slide.
let comments = ppt.Slides.get_Item(0).Comments;

//Save the comments in txt file.
for (let i = 0; i < comments.length; i++){
    str.push(comments[i].Text + "\r\n");
}

//Save to file.
FS.writeFile(outputFileName, str.join(""));

// Clean up resources
ppt.Dispose();
```

---

# spire.presentation javascript comments
## get slide comments from powerpoint
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load document from disk
ppt.LoadFromFile("input.pptx");

//Loop through comments
for (let i = 0; i < ppt.CommentAuthors.Count; i++) {
    let commentAuthor = ppt.CommentAuthors.get_Item(i);
    for (let j = 0; j < commentAuthor.CommentsList.Count; j++) {
        let comment = commentAuthor.CommentsList.get_Item(j);
        //Get comment information
        let commentText = comment.Text;
        let authorName = comment.AuthorName;
        let time = comment.DateTime;
    }
}

// Clean up resources
ppt.Dispose();
```

---

# spire.presentation javascript ppt to svg
## convert PowerPoint slides to SVG while retaining notes
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the file from disk.
ppt.LoadFromFile(inputFileName);

//Retain the notes while converting PowerPoint file to svg file.
ppt.IsNoteRetained = true;
let outFileName = "";
for (let i = 0; i < ppt.Slides.Count; i++) {
    //Convert presentation slides to svg file.
    let bytes = ppt.Slides.get_Item(i).SaveToSVG();
    outFileName = `output_${i}.svg`;          
    bytes.Save(outFileName);
}
```

---

# spire.presentation javascript slide notes
## remove note from specific slide
```javascript
//Get the first slide
let slide = ppt.Slides.get_Item(0);

//Get note slide
let note = slide.NotesSlide;
//Clear note text
note.NotesTextFrame.Text = "";
```

---

# Remove Speaker Notes from PowerPoint Slide
## This code demonstrates how to remove speaker notes from a presentation slide using JavaScript
```javascript
//Get the first slide from the sample document.
let slide = ppt.Slides.get_Item(0);

//Remove the first speak note.
slide.NotesSlide.NotesTextFrame.Paragraphs.RemoveAt(1);
```

---

# Spire.Presentation JavaScript Comment Reply
## Add and delete replies to comments in a presentation
```javascript
//Create Comment author
let author = ppt.CommentAuthors.AddAuthor("E-iceblue", "comment");

//Add comment
ppt.Slides.get_Item(0).AddComment({author:author, text:"Add comment", position:wasmModule.PointF.Create(18, 25),dateTime: wasmModule.DateTime.get_Now()});
let comment = ppt.Slides.get_Item(0).Comments[0];

//Add reply to Comment
if (!comment.IsReply) {
    comment.Reply(author, "Add Reply1", wasmModule.DateTime.get_Now());
    comment.Reply(author, "Add Reply2", wasmModule.DateTime.get_Now());
}

//delete first reply
ppt.Slides.get_Item(0).DeleteComment({author:author, text:"Add Reply1"});
```

---

# spire presentation header footer
## add footer to powerpoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

ppt.LoadFromFile(inputFileName);

//Add footer
ppt.SetFooterText("Demo of Spire.Presentation");

//Set the footer visible
ppt.FooterVisible = true;

//Set the page number visible
ppt.SlideNumberVisible = true;

//Set the date visible
ppt.DateTimeVisible = true;
```

---

# spire.presentation javascript header footer
## manage note master header and footer
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load presentation
ppt.LoadFromFile(inputFileName);

//Set the note Masters header and footer
let noteMasterSlide = ppt.NotesMaster;
if (noteMasterSlide !== null) {
    for (let i = 0; i < noteMasterSlide.Shapes.Count; i++) {
        let shape =  noteMasterSlide.Shapes.get_Item(i);
        if (shape.Placeholder !== null) {
            if (shape.Placeholder.Type == wasmModule.PlaceholderType.Header) {
                shape.TextFrame.Text = "change the header by Spire";
            }
            if (shape.Placeholder.Type == wasmModule.PlaceholderType.Footer) {
                shape.TextFrame.Text = "change the footer by Spire";
            }
        }
    }
}
```

---

# SmartArt Child Nodes Access
## Access and read parameters of SmartArt child nodes in a PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the PPT
ppt.LoadFromFile(inputFileName);

let strB = [];
strB.push('Access SmartArt child nodes.');
strB.push('Here is the SmartArt child node parameters details:');
let outString = '';
for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
  let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
  if (shape instanceof wasmModule.ISmartArt) {
    //Get the SmartArt and collect nodes
    let sa = shape;
    let nodes = sa.Nodes;

    let position = 0;
    //Access the parent node at position 0
    let node = nodes.get_Item(position);
    let childnode;
    //Traverse through all child nodes inside SmartArt
    for (let j = 0; j < node.ChildNodes.Count; j++) {
      //Access SmartArt child node at index i
      childnode = node.ChildNodes.get_Item(j);
      //Print the SmartArt child node parameters
      outString = `Node text = ${childnode.TextFrame.Text}, Node level = ${childnode.Level}, Node Position = ${childnode.Position}`;
      strB.push(outString);
    }
  }
}
```

---

# Accessing SmartArt Nodes
## This code demonstrates how to access SmartArt nodes in a PowerPoint presentation and extract their parameters.
```javascript
for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
  let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
  if (shape instanceof wasmModule.ISmartArt) {
    //Get the SmartArt and collect nodes
    let sa = shape;
    let nodes = sa.Nodes;

    //Traverse through all nodes inside SmartArt
    for (let j = 0; j < nodes.Count; j++) {
      //Access SmartArt node at index j
      let node = nodes.get_Item(j);
      //Get the SmartArt node parameters
      let nodeText = node.TextFrame.Text;
      let nodeLevel = node.Level;
      let nodePosition = node.Position;
    }
  }
}
```

---

# Access SmartArt Layout in PowerPoint Presentation
## This code demonstrates how to access SmartArt layout information from a PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the PPT
ppt.LoadFromFile(inputFileName);

let strB = [];

for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
  let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
  if (shape instanceof wasmModule.ISmartArt) {
    //Get the SmartArt and collect nodes
    let sa = shape;
    //Check SmartArt Layout
    let layout = sa.LayoutType.toString();
    strB.push('SmartArt layout type is ' + layout);
  }
}
```

---

# SmartArt Child Node Access
## Access specific SmartArt child node and retrieve its properties
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the PPT
ppt.LoadFromFile(inputFileName);

for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
  let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
  if (shape instanceof wasmModule.ISmartArt) {
    //Get the SmartArt and collect nodes
    let sa = shape;
    //Get SmartArt node collection
    let nodes = sa.Nodes;

    //Access SmartArt node at index 0
    let node = nodes.get_Item(0);

    //Access SmartArt child node at index 1
    let childNode = node.ChildNodes.get_Item(1);

    //Print the SmartArt child node parameters
    let outString = `Node text = ${childNode.TextFrame.Text}, Node level = ${childNode.Level}, Node Position = ${childNode.Position}`;
  }
}

// Clean up resources
ppt.Dispose();
```

---

# SmartArt Node Management
## Add SmartArt nodes by position in PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load the PPT
ppt.LoadFromFile(inputFileName);

for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
  let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
  if (shape instanceof wasmModule.ISmartArt) {
    // Get the SmartArt and collect nodes
    let smartArt = shape;
    let position = 0;
    // Add a new node at specific position
    let node = smartArt.Nodes.AddNodeByPosition(position);
    // Add text and set the text style
    node.TextFrame.Text = 'New Node';
    node.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
    node.TextFrame.TextRange.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Red;

    // Get a node
    node = smartArt.Nodes.get_Item(1);
    position = 1;
    // Add a new child node at specific position
    let childNode = node.ChildNodes.AddNodeByPosition(position);
    // Add text and set the text style
    node.TextFrame.Text = 'New child node';
    node.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
    node.TextFrame.TextRange.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Blue;
  }
}
```

---

# Spire.Presentation JavaScript SmartArt
## Add a new node to SmartArt in PowerPoint presentation
```javascript
//Get the SmartArt
let sa = ppt.Slides.get_Item(0).Shapes.get_Item(0);

//Add a node
let node = sa.Nodes.AddNode();
//Add text and set the text style
node.TextFrame.Text = 'AddText';
node.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
node.TextFrame.TextRange.Fill.SolidColor.KnownColor = wasmModule.KnownColors.HotPink;
```

---

# SmartArt Assistant Node
## Set SmartArt nodes as assistant nodes in PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load the PPT
ppt.LoadFromFile(inputFileName);

let node;
for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
  let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
  if (shape instanceof wasmModule.ISmartArt) {
    // Get the SmartArt and collect nodes
    let smartArt = shape;
    let nodes = smartArt.Nodes;

    // Traverse through all nodes inside SmartArt
    for (let j = 0; j < nodes.Count; j++) {
      // Access SmartArt node at index i
      node = nodes.get_Item(j);
      // Check if node is assistant node
      if (!node.IsAssistant) {
        // Set node as assistant node
        node.IsAssistant = true;
      }
    }
  }
}
```

---

# spire.presentation javascript smartart
## change text of smartart node
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load the PPT
ppt.LoadFromFile(inputFileName);

for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
  let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
  if (shape instanceof wasmModule.ISmartArt) {
    // Get the SmartArt and collect nodes
    let smartArt = shape;
    // Obtain the reference of a node by using its Index
    // select second root node
    let node = smartArt.Nodes.get_Item(1);
    // Set the text of the TextFrame
    node.TextFrame.Text = 'Second root node';
  }
}
```

---

# spire.presentation javascript smartart
## change SmartArt color style in PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the PPT
ppt.LoadFromFile(inputFileName);

for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
  let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
  if (shape instanceof wasmModule.ISmartArt) {
    //Get the SmartArt and collect nodes
    let smartArt = shape;
    // Check SmartArt color type
    if (smartArt.ColorStyle == wasmModule.SmartArtColorType.ColoredFillAccent1) {
      smartArt.ColorStyle = wasmModule.SmartArtColorType.ColorfulAccentColors;
    }
  }
}
```

---

# spire.presentation javascript smartart
## change smartart shape style
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the PPT
ppt.LoadFromFile(inputFileName);

for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
  let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
  if (shape instanceof wasmModule.ISmartArt) {
    //Get the SmartArt and collect nodes
    let smartArt = shape;
    //Check SmartArt style
    if (smartArt.Style == wasmModule.SmartArtStyleType.SimpleFill) {
      //Change SmartArt Style
      smartArt.Style = wasmModule.SmartArtStyleType.Cartoon;
    }
  }
}
```

---

# Spire.Presentation JavaScript SmartArt
## Create and configure SmartArt shapes in PowerPoint presentations
```javascript
let sa = ppt.Slides.get_Item(0).Shapes.AppendSmartArt(200, 60, 300, 300, wasmModule.SmartArtLayoutType.Gear);

//Set type and color of smartart
sa.Style = wasmModule.SmartArtStyleType.SubtleEffect;
sa.ColorStyle = wasmModule.SmartArtColorType.GradientLoopAccent3;

//Remove all shapes
for (let i = sa.Nodes.Count - 1; i >= 0; i--) {
  sa.Nodes.RemoveNode({ index: i });
}

//Add two custom shapes with text
let node = sa.Nodes.AddNode();
sa.Nodes.get_Item(0).TextFrame.Text = 'aa';
node = sa.Nodes.AddNode();
node.TextFrame.Text = 'bb';
node.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
node.TextFrame.TextRange.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Black;
```

---

# spire.presentation javascript smartart
## extract text from smartart in powerpoint
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the PPT
ppt.LoadFromFile(inputFileName);

//Traverse through all the slides of the PPT file and find the SmartArt shapes.
let st = [];
st.push('Below is extracted text from SmartArt:');
for (let i = 0; i < ppt.Slides.Count; i++) {
  for (let j = 0; j < ppt.Slides.get_Item(i).Shapes.Count; j++) {
    if (ppt.Slides.get_Item(i).Shapes.get_Item(j) instanceof wasmModule.ISmartArt) {
      let smartArt = ppt.Slides.get_Item(i).Shapes.get_Item(j);

      //Extract text from SmartArt and append to the StringBuilder object.
      for (let k = 0; k < smartArt.Nodes.Count; k++) {
        st.push(smartArt.Nodes.get_Item(k).TextFrame.Text);
      }
    }
  }
}
```

---

# spire.presentation javascript smartart
## remove smartart node at specific position
```javascript
//Get the SmartArt and collect nodes
let sa = ppt.Slides.get_Item(0).Shapes.get_Item(0);
let nodes = sa.Nodes;

//Remove the node to specific position
nodes.RemoveNodeByPosition(2);
```

---

# spire.presentation javascript smartart
## set smartart linkline outline
```javascript
let smartArt = ppt.Slides.get_Item(0).Shapes.get_Item(0);
let count = smartArt.Nodes.Count;
let node;
//Loop through all smartArts
for (let i = 0; i < count; i++) {
  node = smartArt.Nodes.get_Item(i);
  //Set the line type
  node.LinkLine.FillType = wasmModule.FillFormatType.Solid;
  //Set the line color
  node.LinkLine.SolidFillColor.Color = wasmModule.Color.get_Red();
  //Set the line width
  node.LinkLine.Width = 2;
  //Set the line DashStyle
  node.LinkLine.DashStyle = wasmModule.LineDashStyleType.SystemDash;
}
```

---

# Spire.Presentation JavaScript SmartArt
## Set SmartArt node outline properties
```javascript
// Get the SmartArt shape from the first slide
let smartArt = ppt.Slides.get_Item(0).Shapes.get_Item(0);
let count = smartArt.Nodes.Count;
let node;

// Loop through all nodes
for (let i = 0; i < count; i++) {
  node = smartArt.Nodes.get_Item(i);
  // Set the fill format type
  node.Line.FillType = wasmModule.FillFormatType.Solid;
  // Set the line color
  node.Line.SolidFillColor.Color = wasmModule.Color.get_Red();
  // Set the line width
  node.Line.Width = 2;
  // Set the line style
  node.Line.Style = wasmModule.TextLineStyle.ThinThin;
}
```

---

# spire.presentation javascript watermark
## add image watermark to PowerPoint slide
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the PPT
ppt.LoadFromFile(inputFileName);

let stream = wasmModule.Stream.CreateByFile(imageFileName);
let image = ppt.Images.Append({ stream: stream });
stream.Close();

//Set the properties of SlideBackground, and then fill the image as watermark.
ppt.Slides.get_Item(0).SlideBackground.Type = wasmModule.BackgroundType.Custom;
ppt.Slides.get_Item(0).SlideBackground.Fill.FillType = wasmModule.FillFormatType.Picture;
ppt.Slides.get_Item(0).SlideBackground.Fill.PictureFill.FillType = wasmModule.PictureFillType.Stretch;
ppt.Slides.get_Item(0).SlideBackground.Fill.PictureFill.Picture.EmbedImage = image;
```

---

# Adding Watermark to PowerPoint Presentation
## This code demonstrates how to add a text watermark to a PowerPoint slide
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the PPT
ppt.LoadFromFile(inputFileName);

//Define a rectangle range
let left = (ppt.SlideSize.Size.Width - 400) / 2;
let top = (ppt.SlideSize.Size.Height - 300) / 2;
let rect = wasmModule.RectangleF.FromLTRB(left, top, 400 + left, 300 + top);

//Add a rectangle shape with a defined range
let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({
  shapeType: wasmModule.ShapeType.Rectangle,
  rectangle: rect,
});

//Set the style of the shape
shape.Fill.FillType = wasmModule.FillFormatType.None;
shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
shape.Rotation = -45;
shape.Locking.SelectionProtection = true;
shape.Line.FillType = wasmModule.FillFormatType.None;

//Add text to the shape
shape.TextFrame.Text = 'E-iceblue';
let textRange = shape.TextFrame.TextRange;
//Set the style of the text range
textRange.Fill.FillType = wasmModule.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = wasmModule.Color.FromArgb({
  alpha: 120,
  baseColor: wasmModule.Color.get_HotPink(),
});
textRange.FontHeight = 50;
```

---

# spire.presentation javascript watermark
## remove text and image watermarks from presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the PPT
ppt.LoadFromFile(inputFileName);

//Remove text watermark by removing the shape which contains the text string "E-iceblue".
for (let i = 0; i < ppt.Slides.Count; i++) {
  for (let j = 0; j < ppt.Slides.get_Item(i).Shapes.Count; j++) {
    if (ppt.Slides.get_Item(i).Shapes.get_Item(j) instanceof wasmModule.IAutoShape) {
      let shape = ppt.Slides.get_Item(i).Shapes.get_Item(j);
      if (shape.TextFrame.Text.includes('E-iceblue')) {
        ppt.Slides.get_Item(i).Shapes.Remove(shape);
      }
    }
  }
}

//Remove image watermark.
for (let i = 0; i < ppt.Slides.Count; i++) {
  ppt.Slides.get_Item(i).SlideBackground.Fill.FillType = wasmModule.FillFormatType.None;
}
```

---

# Spire.Presentation JavaScript OLE
## Embed Excel as OLE object in PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

let stream = wasmModule.Stream.CreateByFile(imageFileName);
let oleImage = ppt.Images.Append({ stream });
stream.Close();

let rec = wasmModule.RectangleF.FromLTRB(80, 60, 550, 450);
//Insert an OLE object to presentation based on the Excel data
let objectData = wasmModule.Stream.CreateByFile(inputFileName);
let oleObject = ppt.Slides.get_Item(0).Shapes._AppendOleObject('excel', objectData, rec);
oleObject.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage;

oleObject.ProgId = 'Excel.Sheet.12';
```

---

# spire.presentation javascript OLE embedding
## embed zip file into PowerPoint as OLE object
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();
//Load the PPT
ppt.LoadFromFile(inputFileName);

//Load a zip object
let data = wasmModule.Stream.CreateByFile(inputFile_zName);

let rec = wasmModule.RectangleF.FromLTRB(80, 60, 180, 160);

//Insert the zip object to presentation
let ole = ppt.Slides.get_Item(0).Shapes._AppendOleObjectOOR(inputFile_zName, data, rec);
ole.ProgId = 'Package';
let stream = wasmModule.Stream.CreateByFile(inputFile_IName);
let oleImage = ppt.Images.Append({ stream: stream });

ole.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage;
```

---

# spire.presentation javascript OLE
## extract OLE objects from PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the PPT
ppt.LoadFromFile(inputFileName);

//Loop through the slides and shapes
for (let i = 0; i < ppt.Slides.Count; i++) {
  let slide = ppt.Slides.get_Item(i);
  for (let j = 0; j < slide.Shapes.Count; j++) {
    let shape = slide.Shapes.get_Item(j);
    if (shape instanceof wasmModule.IOleObject) {
      //Find OLE object
      let oleObject = shape;

      //Get its data
      let bytes = oleObject.Data;
      // Determine the type of OLE object by its ProgId
      let oleType = oleObject.ProgId;
    }
  }
}

// Clean up resources
ppt.Dispose();
```

---

# Spire.Presentation JavaScript OLE
## Modify OLE object data in PowerPoint presentation
```javascript
// Create a PPT document
let ppt = wasmModule.Presentation.Create();

// Load the PPT
ppt.LoadFromFile(inputFileName);

// Loop through the slides and shapes
for (let i = 0; i < ppt.Slides.Count; i++) {
  let slide = ppt.Slides.get_Item(i);
  for (let j = 0; j < slide.Shapes.Count; j++) {
    let shape = slide.Shapes.get_Item(j);
    if (shape instanceof wasmModule.IOleObject) {
      // Find OLE object
      let oleObject = shape;

      // Get its data and write to file
      let bytes = oleObject.Data;
      let pptStream = wasmModule.Stream.CreateByBytes(bytes);
      let stream = wasmModule.Stream.Create();
      if (oleObject.ProgId == 'PowerPoint.Show.12') {
        // Load the PPT stream
        let ppt1 = wasmModule.Presentation.Create();
        ppt1.LoadFromStream({ stream: pptStream, fileFormat: wasmModule.FileFormat.Auto });
        ppt1.Slides.get_Item(0).Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: imageFileName, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 150, 150) });
        ppt1.SaveToFile({ stream: stream, fileFormat: wasmModule.FileFormat.Pptx2013 });
        stream.Position = BigInt(0);
        // Modify the data
        oleObject.Data = stream;
      }
    }
  }
}

// Save result file
ppt.SaveToFile({
  file: outputFileName,
  fileFormat: wasmModule.FileFormat.Pptx2013,
});
```

---

# spire.presentation javascript vba
## remove VBA macros from PowerPoint presentation
```javascript
//Create a PPT document
let ppt = wasmModule.Presentation.Create();

//Load the PPT
ppt.LoadFromFile(inputFileName);

//Remove macros
ppt.DeleteMacros();

// Save result file
ppt.SaveToFile({
  file: outputFileName,
  fileFormat: wasmModule.FileFormat.Pptx2013,
});

// Clean up resources
ppt.Dispose();
```

---

