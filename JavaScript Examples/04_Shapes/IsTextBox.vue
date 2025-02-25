<template>
  <span>The example demonstrates how to detect if the shape is textbox.</span>
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

        const inputFileName = "IsTextboxSample.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load file
        ppt.LoadFromFile(inputFileName);

        let stringBuilder = [];

        for (let i = 0; i < ppt.Slides.Count; i++) {
            let slide = ppt.Slides.get_Item(i);
            for (let j = 0; j < slide.Shapes.Count; j++) {
                let shape = slide.Shapes.get_Item(j);
                if(shape instanceof wasmModule.IAutoShape){
                    //Judge if the shape is textbox
                    let isTextbox = shape.IsTextBox;
                    stringBuilder.push(isTextbox ? "shape is text box" : "shape is not text box")
                }
            }
        }

        // Define the output file name
        const outputFileName = "IsTextBox_out.txt";
        wasmModule.FS.writeFile(outputFileName, stringBuilder.join("\n"));

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
