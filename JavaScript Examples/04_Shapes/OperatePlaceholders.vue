<template>
  <span>The example demonstrates how to add image/chart/table/smartArt to the placeholders in a PPT document. </span>
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

        const inputFileName = "OperatePlaceholders.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        const videoFileName = "Video.mp4";
        await wasmModule.FetchFileToVFS(videoFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        const imageFileName = "E-iceblueLogo.png";
        await wasmModule.FetchFileToVFS(imageFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the document from disk
        ppt.LoadFromFile(inputFileName);

        //Operate placeholders
        for (let j = 0; j < ppt.Slides.Count; j++){
            let slide = ppt.Slides.get_Item(j);
            for (let i = 0; i < slide.Shapes.Count; i++) {
                let shape = slide.Shapes.get_Item(i);
                switch (shape.Placeholder.Type) {
                    case wasmModule.PlaceholderType.Media:
                        shape.InsertVideo(videoFileName);
                        break;
                    case wasmModule.PlaceholderType.Picture:
                        shape.InsertPicture({filepath:imageFileName});
                        break;

                    case wasmModule.PlaceholderType.Chart:
                        shape.InsertChart(wasmModule.ChartType.ColumnClustered);
                        break;

                    case wasmModule.PlaceholderType.Table:
                        shape.InsertTable(3, 2);
                        break;

                    case wasmModule.PlaceholderType.Diagram:
                        shape.InsertSmartArt(wasmModule.SmartArtLayoutType.BasicBlockList);
                        break;
                }
            }
        }
        // Define the output file name
        const outputFileName = "OperatePlaceholders_out.pptx";

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
