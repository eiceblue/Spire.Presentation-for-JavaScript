<template>
  <span>Click the following button to add superscript and subscript.</span>
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
        let inputFileName = "Template_Az.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName); 

        // Get the first slide
        let slide = ppt.Slides.get_Item(0);
        
        // Append a rectangle shape to the slide
        let shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(150, 100, 350, 150)});
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
        shape.Fill.FillType = wasmModule.FillFormatType.None;
        shape.TextFrame.Paragraphs.Clear();

        // Append a text frame with the initial text "Test".
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

        const outputFileName = "SuperscriptAndSubscript.pptx";

        // Save to file
        ppt.SaveToFile({file:outputFileName,fileFormat:wasmModule.FileFormat.Pptx2013});

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
