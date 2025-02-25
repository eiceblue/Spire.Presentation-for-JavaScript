<template>
  <span>The example shows how to get the shapes by placeholder.</span>
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

        const inputFileName = "GetShapesByPlaceholder.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();
        ppt.LoadFromFile(inputFileName);

        let placeholder = ppt.Slides.get_Item(1).Shapes.get_Item(0).Placeholder;
        //Get Shapes by Placeholder
        let shapes = ppt.Slides.get_Item(1).GetPlaceholderShapes(placeholder);

        let text = "";
        //Iterate over all the shapes
        for (let i = 0; i < shapes.length; i++){
            //If shape is IAutoShape
            if (shapes[i] instanceof wasmModule.IAutoShape){
                let autoShape = shapes[i];
                if (autoShape.TextFrame != null) {
                    text += autoShape.TextFrame.Text + "\r\n";
                }
            }
        }

        // Define the output file name
        const outputFileName = "GetShapesByPlaceholder_out.txt";
        wasmModule.FS.writeFile(outputFileName, text);

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "text/plain"});

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
