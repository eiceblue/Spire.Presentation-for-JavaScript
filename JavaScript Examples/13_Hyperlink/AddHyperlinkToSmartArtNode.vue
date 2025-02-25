<template>
  <span>Click the following button to add hyperlink to SmartArt nodes in a PPT document.</span>
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

        let inputFileName = "SmartArtNode.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);

        //Get the smartArt shape
        let sr = ppt.Slides.get_Item(0).Shapes.get_Item(0);
        //Add hylerlinks to the nodes
        let node = sr.Nodes.get_Item(0);
        node.Click = wasmModule.ClickHyperlink.Create_silde(ppt.Slides.get_Item(1));
        node = sr.Nodes.get_Item(1);
        node.Click = wasmModule.ClickHyperlink.Create_silde(ppt.Slides.get_Item(2));
        node = sr.Nodes.get_Item(2);
        node.Click = wasmModule.ClickHyperlink.Create_silde(ppt.Slides.get_Item(3));


        // Define the output file name
        const outputFileName = "AddHyperlinkToSmartArtNode_out.pptx";

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
