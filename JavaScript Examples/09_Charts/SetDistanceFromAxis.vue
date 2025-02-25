<template>
    <span>The following example demonstrates how to set distance from axis for chart in a PPT document</span>
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

                // Create a ppt document
                let ppt = wasmModule.Presentation.Create();

                // Append ColumnClustered chart
                let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.ColumnClustered, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 450, 450) });

                // Get the PrimaryCategory axis
                let chartAxis = chart.PrimaryCategoryAxis;

                // Set "Distance from axis"
                chartAxis.LabelsDistance = 200;

                // Save to file
                const outputFileName = "SetDistanceFromAxis.pptx";
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