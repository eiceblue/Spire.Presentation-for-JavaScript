<template>
  <span>Click the following button to create slides.</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let ImageFileName = "bg.png";
        await wasmModule.FetchFileToVFS(ImageFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Add new slide
        ppt.Slides.Append();

        // Set the background image
        for (let i = 0; i < 2; i++) {
          let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
          ppt.Slides.get_Item(i).Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: ImageFileName, rectangle: rect });
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

        //Add new paragraph
        let pare = wasmModule.TextParagraph.Create();
        pare.Text = "";
        shape.TextFrame.Paragraphs._Append(pare);

        //Add new paragraph
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

        const outputFileName = "CreateSlide.pptx";

        // Save to file
        ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });

        // Clean up resources
        ppt.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
