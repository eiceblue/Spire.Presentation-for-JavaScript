<template>
  <span>Click the following button to modify OLE object data in a PPT document. </span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName"> Click here to download the generated file </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref('');

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = 'ModifyOLEData.pptx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);
        let imageFileName = 'Logo.png';
        await wasmModule.FetchFileToVFS(imageFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the PPT
        ppt.LoadFromFile(inputFileName);

        //Loop through the slides and shapes
        for (let i = 0; i < ppt.Slides.Count; i++) {
          let slide = ppt.Slides.get_Item(i);
          for (let j = 0; j < slide.Shapes.Count; j++) {
            let shape = slide.Shapes.get_Item(j);
            if (shape instanceof wasmModule.IOleObject) {
              //Find OLE object
              let oleObject = shape;

              //Get its data and write to file
              let bytes = oleObject.Data;
              let pptStream = wasmModule.Stream.CreateByBytes(bytes);
              let stream = wasmModule.Stream.Create();
              if (oleObject.ProgId == 'PowerPoint.Show.12') {
                //Load the PPT stream

                let ppt1 = wasmModule.Presentation.Create();
                ppt1.LoadFromStream({ stream: pptStream, fileFormat: wasmModule.FileFormat.Auto });
                ppt1.Slides.get_Item(0).Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: imageFileName, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 150, 150) });
                ppt1.SaveToFile({ stream: stream, fileFormat: wasmModule.FileFormat.Pptx2013 });
                stream.Position = BigInt(0);
                //Modify the data
                oleObject.Data = stream;
              }
            }
          }
        }
        // Define the output file name
        const outputFileName = 'ModifyOLEData.pptx';

        // Save result file
        ppt.SaveToFile({
          file: outputFileName,
          fileFormat: wasmModule.FileFormat.Pptx2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        });

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
