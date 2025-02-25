<template>
  <span>The sample demonstrates how to set the background for a PPT slide. </span>
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

        const ImageFileName = "backgroundImg.png";
        await wasmModule.FetchFileToVFS(ImageFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

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

        // Define the output file name
        const outputFileName = "Background_out.pptx";

        // Save the document to the specified path
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
