<template>
  <span>The example shows how to reset the position of date time and slide number placeholder. </span>
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

        const inputFileName = "Template_Ppt_7.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the file from disk.
        ppt.LoadFromFile(inputFileName);

        //Get the first slide from the sample document.
        let slide = ppt.Slides.get_Item(0);

        for (let i = 0; i < slide.Shapes.Count; i++) {
            let shape = slide.Shapes.get_Item(i);
            //Reset the position of the slide number to the left.
            if(shape.Name.includes("Slide Number Placeholder")){
                shape.Left = 0;
            }else if(shape.Name.includes("Date Placeholder")){
                //Reset the position of the date time to the center.
                shape.Left = ppt.SlideSize.Size.Width / 2;
                //Reset the date time display style.
                shape.TextFrame.TextRange.Paragraph.Text = wasmModule.DateTime.get_Now().ToString({format:"dd.MM.yyyy"});
                shape.TextFrame.IsCentered = true;
            }
        }

        // Define the output file name
        const outputFileName = "ResetPositionOfPlaceholder_out.pptx";

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
