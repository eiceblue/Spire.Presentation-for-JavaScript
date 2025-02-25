<template>
  <span>The following example demonstrates how to add image into table cell</span>
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
         
        // Load the input file and image into the virtual file system (VFS)
        const inputFileName = "AddImageInTableCell.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);
        const fileStreamName = "PresentationIcon.png";
        await wasmModule.FetchFileToVFS(fileStreamName, "", `${import.meta.env.BASE_URL}static/data/`);
        
        // Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();
        ppt.LoadFromFile(inputFileName);

        //Get the first shape
        let table = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        //Load the image and insert it into table cell
        let stream = wasmModule.Stream.CreateByFile(fileStreamName);
        let pptImg = ppt.Images.Append({ stream: stream });
        stream.Close();

        table.get_Item(1, 1).FillFormat.FillType = wasmModule.FillFormatType.Picture;
        table.get_Item(1, 1).FillFormat.PictureFill.Picture.EmbedImage = pptImg;
        table.get_Item(1, 1).FillFormat.PictureFill.FillType = wasmModule.PictureFillType.Stretch;

        // Define the output file name
        const outputFileName = "AddImageInTableCell.pptx";

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
