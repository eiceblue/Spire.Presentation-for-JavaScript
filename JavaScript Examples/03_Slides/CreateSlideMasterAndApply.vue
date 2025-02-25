<template>
  <span>Click the following button to create multiple slide masters and apply them in a PPT.</span>
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

        // Load the sample images into the virtual file system (VFS)
        let pic1Name = "bg.png";
        await wasmModule.FetchFileToVFS(pic1Name, "", `${import.meta.env.BASE_URL}static/data/`);

        let pic2Name = "Setbackground.png";
        await wasmModule.FetchFileToVFS(pic2Name, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        ppt.SlideSize.Type = wasmModule.SlideSizeType.Screen16x9;

        // Add slides
        for (let i = 0; i < 4; i++) {
          ppt.Slides.Append();
        }

        // Get the first default slide master
        let first_master = ppt.Masters.get_Item(0);

        // Append another slide master
        ppt.Masters.AppendSlide(first_master);
        let second_master = ppt.Masters.get_Item(1);

        // The first slide masters
        let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        first_master.SlideBackground.Fill.FillType = wasmModule.FillFormatType.Picture;
        let image1 = first_master.Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: pic1Name, rectangle: rect });
        first_master.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image1.PictureFill.Picture.EmbedImage;
        
        // The second slide master
        second_master.SlideBackground.Fill.FillType = wasmModule.FillFormatType.Picture;
        let image2 = second_master.Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: pic2Name, rectangle: rect });
        second_master.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image2.PictureFill.Picture.EmbedImage;

        // Apply the first master with layout to the first slide
        ppt.Slides.get_Item(0).Layout = first_master.Layouts.get_Item(1);

        // Apply the second master with layout to other slides
        for (let i = 1; i < ppt.Slides.Count; i++) {
          ppt.Slides.get_Item(i).Layout = second_master.Layouts.get_Item(8);
        }

        const outputFileName = "CreateSlideMasterAndApply.pptx";

        // Save to file
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
