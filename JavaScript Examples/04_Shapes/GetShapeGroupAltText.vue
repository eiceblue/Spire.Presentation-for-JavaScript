<template>
  <span>The example demonstrates how to get alt text of shapes in shape group.</span>
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

        const inputFileName = "GetShapeGroupAltText.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load document from disk
        ppt.LoadFromFile(inputFileName);

        let builder = [];

        //Loop through slides and shapes
        for (let i = 0;i < ppt.Slides.Count;i++) {
            let slide = ppt.Slides.get_Item(i);
            for (let j = 0; j < slide.Shapes.Count; j++) {
                let shape = slide.Shapes.get_Item(j);
                if(shape instanceof wasmModule.GroupShape){
                    //Find the shape group
                    let groupShape = shape;
                    for (let k = 0;k < groupShape.Shapes.Count;k++){
                        let gShape = groupShape.Shapes.get_Item(k);
                        //Append the alternative text in builder
                        builder.push(gShape.AlternativeText);
                    }
                }
            }
        }
        // Define the output file name
        const outputFileName = "GetShapeGroupAltText_out.txt";
        wasmModule.FS.writeFile(outputFileName, builder.join("\n"));

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
