<template>
  <span>Click the following button to append HTML into PPT document.</span>
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
        let inputFileName = "AppendHTML.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

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

        // Define the output file name 
        const outputFileName = "AppendHTML.pptx";

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
