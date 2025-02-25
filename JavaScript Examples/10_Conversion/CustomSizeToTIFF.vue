<template>
  <span>The following example demonstrates how to convert PPT to TIFF with customized size</span>
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
         
        // Load the input file into the virtual file system (VFS)
        const inputFileName = "Indent.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create PPT document
        let ppt =wasmModule.Presentation.Create();

        // Load the original PPT document from VFS
        ppt.LoadFromFile(inputFileName);

        // Get the first slide
        let slide = ppt.Slides.get_Item(0);

        // Create a new PPT document
        let newPpt =wasmModule.Presentation.Create();

        // Remove the default slide
        newPpt.Slides.RemoveAt(0);

        // Define a new size
        let size =wasmModule.SizeF.CreateWH(200, 200);

        // Set PPT slide size
        newPpt.SlideSize.Size = size;

        // Insert the slide of original PPT
        newPpt.Slides.Insert({ index: 0, slide: slide });


        // Define the output file name
        const outputFileName = "Output1.tiff";

        // Save the document 
        newPpt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Tiff });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/tiff" });

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
