<template>
  <span>The following example demonstrates how to set image transparency in a PPT document</span>
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
        const imagePathName = "iceblueLogo.png";
        await wasmModule.FetchFileToVFS(imagePathName, "", `${import.meta.env.BASE_URL}static/data/`);

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


        // Define the output file name
        const outputFileName = "SetImageTransparency.pptx";

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
