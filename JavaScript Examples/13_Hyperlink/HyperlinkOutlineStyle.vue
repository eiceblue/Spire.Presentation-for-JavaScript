<template>
  <span>Click the following button to add hyperlink and set its outline style in a PPT.</span>
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

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();


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


        // Define the output file name
        const outputFileName = "HyperlinkOutlineStyle_out.pptx";

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
