<template>
  <span>Click the following button to link to a specific slide in PowerPoint document.</span>
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


        // Define the output file name
        const outputFileName = "LinkToASpecificSlide_out.pptx";

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
