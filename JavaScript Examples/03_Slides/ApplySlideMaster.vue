<template>
  <span>Click the following button to apply slide master in a PPT document.</span>
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
        let inputFileName = "InputTemplate.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Load the sample file into the virtual file system (VFS)
        let backgroundPicName = "bg.png";
        await wasmModule.FetchFileToVFS(backgroundPicName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName);

        // Get the first slide master from the presentation
        let masterSlide = ppt.Masters.get_Item(0);

        // Customize the background of the slide master
        let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        masterSlide.SlideBackground.Fill.FillType = wasmModule.FillFormatType.Picture;
        let image = masterSlide.Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: backgroundPicName, rectangle: rect });
        masterSlide.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image.PictureFill.Picture.EmbedImage;

        // Change the color scheme
        masterSlide.Theme.ColorScheme.Accent1.Color = wasmModule.Color.get_Red();
        masterSlide.Theme.ColorScheme.Accent2.Color = wasmModule.Color.get_RosyBrown();
        masterSlide.Theme.ColorScheme.Accent3.Color = wasmModule.Color.get_Ivory();
        masterSlide.Theme.ColorScheme.Accent4.Color = wasmModule.Color.get_Lavender();
        masterSlide.Theme.ColorScheme.Accent5.Color = wasmModule.Color.get_Black();

        const outputFileName = "ApplySlideMaster.pptx";

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
