<template>
  <span>Click the following button to set text format of table.</span>
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

        let inputFileName = "Table.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load PPT file from disk
        ppt.LoadFromFile(inputFileName);
        //Get the first slide
        let slide = ppt.Slides.get_Item(0);
        let str = [];
        for (let i = 0; i < slide.Shapes.Count; i++) {
            let shape = slide.Shapes.get_Item(i);
            //Verify if it is table
            if (shape instanceof wasmModule.ITable){
                let table = shape;
                let cell1 = table.TableRows.get_Item(0).get_Item(0);
                //Set table cell's text alignment type
                cell1.TextAnchorType = wasmModule.TextAnchorType.Top;
                //Set italic style
                cell1.TextFrame.TextRange.Format.IsItalic = wasmModule.TriState.True;

                let cell2 = table.TableRows.get_Item(1).get_Item(0);
                //Set table cell's foreground color
                cell2.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
                cell2.TextFrame.TextRange.Fill.SolidColor.Color = wasmModule.Color.get_Green();
                //Set table cell's background color
                cell2.FillFormat.FillType = wasmModule.FillFormatType.Solid;
                cell2.FillFormat.SolidColor.Color = wasmModule.Color.get_LightGray();


                let cell3 = table.TableRows.get_Item(2).get_Item(2);
                //Set table cell's font and font size
                cell3.TextFrame.TextRange.FontHeight = 12;
                cell3.TextFrame.TextRange.LatinFont = wasmModule.TextFont.Create("Arial Black");
                cell3.TextFrame.TextRange.HighlightColor.Color = wasmModule.Color.get_YellowGreen();


                let cell4 = table.TableRows.get_Item(2).get_Item(1);
                //Set table cell's margin and borders
                cell4.MarginLeft = 20;
                cell4.MarginTop = 30;
                cell4.BorderTop.FillType = wasmModule.FillFormatType.Solid;
                cell4.BorderTop.SolidFillColor.Color = wasmModule.Color.get_Red();
                cell4.BorderBottom.FillType = wasmModule.FillFormatType.Solid;
                cell4.BorderBottom.SolidFillColor.Color = wasmModule.Color.get_Red();
                cell4.BorderLeft.FillType = wasmModule.FillFormatType.Solid;
                cell4.BorderLeft.SolidFillColor.Color = wasmModule.Color.get_Red();
                cell4.BorderRight.FillType = wasmModule.FillFormatType.Solid;
                cell4.BorderRight.SolidFillColor.Color = wasmModule.Color.get_Red();
            }
        }


        // Define the output file name
        const outputFileName = "SetTextFormat_out.pptx";

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
