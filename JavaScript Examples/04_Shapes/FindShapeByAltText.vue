<template>
  <span>The example demonstrates how to find shape by its alt text in a PPT document. </span>
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

        const inputFileName = "FindShapeByAltText.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load document from disk
        ppt.LoadFromFile(inputFileName);

        //Get the first slide
        let slide = ppt.Slides.get_Item(0);

        //Find shape in the slide
        let shape = FindShape(slide, "Shape1");

        let str = [];
        str.push(shape.Name);
        
        // Define the output file name
        const outputFileName = "FindShapeByAltText_out.txt";

        wasmModule.FS.writeFile(outputFileName, str.join("\n"));

        ppt.Dispose();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],{type: "text/plain"});

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
      
function FindShape(slide, altText) {
    //Loop through shapes in the slide
    for (let i = 0;i < slide.Shapes.Count;i++){
        let shape = slide.Shapes.get_Item(i);
        //Find the shape whose alternative text is altText
        if (shape.AlternativeText == altText) {
            return shape;
        }
    }
    return null;
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
