<template>
  <span>The example demonstrates how to prevent or allow changing shape in a PPT document.</span>
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

        const ImageFileName = "bg.png";
        await wasmModule.FetchFileToVFS(ImageFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Set background image
        let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName: ImageFileName,rectangle: rect});
        ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        //Add a rectangle shape to the slide
        let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(50, 100, 450, 250)});

        //Set the shape format
        shape.Fill.FillType = wasmModule.FillFormatType.None;
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_LightBlue();
        shape.TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Justify;
        shape.TextFrame.Text = "Demo for locking shapes:\n    Green/Black stands for editable.\n    Grey stands for non-editable.";
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Arial Rounded MT Bold");
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_Black();

        //The changes of selection and rotation are allowed
        shape.Locking.RotationProtection = false;
        shape.Locking.SelectionProtection = false;
        //The changes of size, position, shape type, aspect ratio, text editing and ajust handles are not allowed
        shape.Locking.ResizeProtection = true;
        shape.Locking.PositionProtection = true;
        shape.Locking.ShapeTypeProtection = true;
        shape.Locking.AspectRatioProtection = true;
        shape.Locking.TextEditingProtection = true;
        shape.Locking.AdjustHandlesProtection = true;

        // Define the output file name
        const outputFileName = "PreventOrAllowChangingShape_out.pptx";
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
