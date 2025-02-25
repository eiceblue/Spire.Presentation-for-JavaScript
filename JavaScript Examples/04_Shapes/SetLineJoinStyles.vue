<template>
  <span>The example demonstrates how to set the join style of shape lines in a PPT document</span>
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

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Get the first slide
        let slide = ppt.Slides.get_Item(0);

        //Add three shapes
        let shape1 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(50, 150, 200, 200)});
        let shape2 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: wasmModule.RectangleF.FromLTRB(250, 150, 400, 200)});
        let shape3 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(450, 150, 600, 200)});

        //Fill shapes
        shape1.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape1.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
        shape2.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape2.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
        shape3.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape3.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();

        //Fill lines of shapes
        shape1.Line.FillType = wasmModule.FillFormatType.Solid;
        shape1.Line.SolidFillColor.Color = wasmModule.Color.get_DarkGray();
        shape2.Line.FillType = wasmModule.FillFormatType.Solid;
        shape2.Line.SolidFillColor.Color = wasmModule.Color.get_DarkGray();
        shape3.Line.FillType = wasmModule.FillFormatType.Solid;
        shape3.Line.SolidFillColor.Color = wasmModule.Color.get_DarkGray();

        //Set the line width
        shape1.Line.Width = 10;
        shape2.Line.Width = 10;
        shape3.Line.Width = 10;

        //Set the join styles of lines
        shape1.Line.JoinStyle = wasmModule.LineJoinType.Bevel;
        shape2.Line.JoinStyle = wasmModule.LineJoinType.Miter;
        shape3.Line.JoinStyle = wasmModule.LineJoinType.Round;

        //Add text in shapes
        shape1.TextFrame.Text = "Bevel Join Style";
        shape2.TextFrame.Text = "Miter Join Style";
        shape3.TextFrame.Text = "Round Join Style";

        // Define the output file name
        const outputFileName = "SetLineJoinStyles_out.pptx";

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
