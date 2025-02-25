<template>
  <span>The following example demonstrates how to set format of image in a PPT document</span>
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

        // Load the image into the virtual file system (VFS)
        const imageFileName = "iceblueLogo.png";
        await wasmModule.FetchFileToVFS(imageFileName, "", `${import.meta.env.BASE_URL}static/data/`);

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

        // Define the output file name
        const outputFileName = "SetImageFrameFormat.pptx";

        // Save the document 
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
