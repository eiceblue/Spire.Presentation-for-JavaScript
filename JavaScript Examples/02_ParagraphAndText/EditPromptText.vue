<template>
  <span>Click the following button to edit the prompt text to the slide.</span>
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

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "HasPromptText.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName);

        // Iterate through the slide
        for (let i = 0;i < ppt.Slides.get_Item(0).Shapes.Count;i++){
            let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
            if(shape.Placeholder != null && shape instanceof wasmModule.IAutoShape){
                let text = "";
                // Set the text of the title
                if (shape.Placeholder.Type == wasmModule.PlaceholderType.CenteredTitle)
                {
                    text = "custom title create by Spire";
                }
                // Set text of the subtitle
                else if (shape.Placeholder.Type == wasmModule.PlaceholderType.Subtitle)
                {
                    text = "custom subtitle create by Spire";
                }
                shape.TextFrame.Text = text;
            }
        }

        const outputFileName = "EditPromptText.pptx";

        // Save to file
        ppt.SaveToFile({file:outputFileName,fileFormat:wasmModule.FileFormat.Pptx2013});

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
