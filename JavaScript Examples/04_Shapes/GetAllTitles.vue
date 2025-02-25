<template>
  <span>The example demonstrates how to get all titles of slides in a PPT document.</span>
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

        const inputFileName = "Titles.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load file
        ppt.LoadFromFile(inputFileName);

        //Instantiate a list of IShape objects
        let shapelist = [];
        //Loop through all sildes and all shapes on each slide
        for (let i = 0;i < ppt.Slides.Count;i++){
            let slide = ppt.Slides.get_Item(i);
            for (let j = 0;j < slide.Shapes.Count;j++){
                let shape = slide.Shapes.get_Item(j);
                if(shape.Placeholder != null){
                    //Get all titles
                    switch (shape.Placeholder.Type) {
                        case wasmModule.PlaceholderType.Title:
                            shapelist.push(shape);
                            break;
                        case wasmModule.PlaceholderType.CenteredTitle:
                            shapelist.push(shape);
                            break;
                        case wasmModule.PlaceholderType.Subtitle:
                            shapelist.push(shape);
                            break;
                    }
                }
            }
        }

        //Loop through the list and get the inner text of all shapes in the list
        let stringBuilder = [];
        stringBuilder.push("Below are all the obtained titles:");
        for (let i = 0; i < shapelist.length; i++){
            let shape1 = shapelist[i];
            stringBuilder.push(shape1.TextFrame.Text);
        }

        // Define the output file name
        const outputFileName = "GetAllTitles_out.txt";
        wasmModule.FS.writeFile(outputFileName, stringBuilder.join("\n"));

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],{type: "text/plain"});

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
