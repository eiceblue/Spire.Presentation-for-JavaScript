<template>
  <span>This sample demonstrates how to insert Table into a PPT document</span>
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
        const inputFileName = "CreateTable.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt =wasmModule.Presentation.Create();

        // Load the document from disk
        ppt.LoadFromFile(inputFileName);

        let widths = [100, 100, 150, 100, 100];
        let heights = [15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15];

        // Add new table to PPT
        let table = ppt.Slides.get_Item(0).Shapes.AppendTable(ppt.SlideSize.Size.Width / 2 - 275, 90, widths, heights);
        let dataStr = [
          ["Name", "Capital", "Continent", "Area", "Population"],
          ["Venezuela", "Caracas", "South America", "912047", "19700000"],
          ["Bolivia", "La Paz", "South America", "1098575", "7300000"],
          ["Brazil", "Brasilia", "South America", "8511196", "150400000"],
          ["Canada", "Ottawa", "North America", "9976147", "26500000"],
          ["Chile", "Santiago", "South America", "756943", "13200000"],
          ["Colombia", "Bagota", "South America", "1138907", "33000000"],
          ["Cuba", "Havana", "North America", "114524", "10600000"],
          ["Ecuador", "Quito", "South America", "455502", "10600000"],
          ["Paraguay", "Asuncion", "South America", "406576", "4660000"],
          ["Peru", "Lima", "South America", "1285215", "21600000"],
          ["Jamaica", "Kingston", "North America", "11424", "2500000"],
          ["Mexico", "Mexico City", "North America", "1967180", "88600000"]
        ];

        // Add data to table
        for (let i = 0; i < 13; i++) {
          for (let j = 0; j < 5; j++) {
            // Fill the table with data
            table.get_Item(j, i).TextFrame.Text = dataStr[i][j];

            // Set the Font
            table.get_Item(j, i).TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).LatinFont =wasmModule.TextFont.Create("Arial Narrow");
          }
        }


        // Set the alignment of the first row to Center
        for (let i = 0; i < 5; i++) {
          table.get_Item(i, 0).TextFrame.Paragraphs.get_Item(0).Alignment =wasmModule.TextAlignmentType.Center;
        }

        // Set the style of table
        table.StylePreset =wasmModule.TableStylePreset.LightStyle3Accent1;

        // Define the output file name
        const outputFileName = "CreateTable.pptx";

        // Save the document 
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
