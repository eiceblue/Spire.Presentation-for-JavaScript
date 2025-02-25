<template>
  <div>
    <el-tree style="max-width: 600px" :data="data" :props="defaultProps" accordion @node-click="handleNodeClick" />
  </div>
</template>
<script setup lang="ts">
interface Tree {
  label: string;
  children?: Tree[];
  $treeNodeId?: string | number;
  [key: string]: any;
}
const emits = defineEmits(["item-id"]);
const handleNodeClick = (data: Tree) => {
  if ((data as any).id !== undefined) {
    emits("item-id", (data as any).id);
  }
};
const data: Tree[] = [
  {
    label: "Spire.Presentation",
    children: [
      // 01_QuickGuide
      {
        label: "01_QuickGuide",
        children: [
          {
            label: "HelloWorld",
            id: "01-1",
          },
        ],
      },
      // 02_ParagraphAndText
      {
        label: "02_ParagraphAndText",
        children: [
          {
            label: "AddParagraph",
            id: "02-1",
          },
          {
            label: "Alignment",
            id: "02-2",
          },
          {
            label: "AppendHTML",
            id: "02-3",
          },
          {
            label: "AutoFitTextOrShape",
            id: "02-4",
          },
          {
            label: "BordersAndShading",
            id: "02-5",
          },
          {
            label: "Bullets",
            id: "02-6",
          },
          {
            label: "ChangeTextStyle",
            id: "02-7",
          },
          {
            label: "CopyParagraphToAnotherPPT",
            id: "02-8",
          },
          {
            label: "CustomBulletsNumber",
            id: "02-9",
          },
          {
            label: "EditPromptText",
            id: "02-10",
          },
          {
            label: "ExtractText",
            id: "02-11",
          },
          {
            label: "GetTextFrameEffectiveData",
            id: "02-12",
          },
          {
            label: "GetTextStyleEffectiveData",
            id: "02-13",
          },
          {
            label: "HighlightSpecifiedText",
            id: "02-14",
          },
          {
            label: "Indent",
            id: "02-15",
          },
          {
            label: "LineSpacing",
            id: "02-16",
          },
          {
            label: "MixFontStyles",
            id: "02-17",
          },
          {
            label: "ModifyStyleOfFirstFoundText",
            id: "02-18",
          },
          {
            label: "MultipleLevelBullets",
            id: "02-19",
          },
          {
            label: "MultipleParagraphs",
            id: "02-20",
          },
          {
            label: "PictureCustomBulletStyle",
            id: "02-21",
          },
          {
            label: "RemoveTextBox",
            id: "02-22",
          },
          {
            label: "ReplaceText",
            id: "02-23",
          },
          {
            label: "ReplaceTextRetentionStyle",
            id: "02-24",
          },
          {
            label: "ReplaceTextWithRegex",
            id: "02-25",
          },
          {
            label: "RotateText",
            id: "02-26",
          },
          {
            label: "Set3DEffectForText",
            id: "02-27",
          },
          {
            label: "SetAnchorOfTextFrame",
            id: "02-28",
          },
          {
            label: "SetColumnsCountOfTextFrame",
            id: "02-29",
          },
          {
            label: "SetCustomFonts",
            id: "02-30",
          },
          {
            label: "SetParagraphFont",
            id: "02-31",
          },
          {
            label: "SetRightToLeftColumns",
            id: "02-32",
          },
          {
            label: "SetShadowEffectForText",
            id: "02-33",
          },
          {
            label: "SetTextDirection",
            id: "02-34",
          },
          {
            label: "SetTextFontProperties",
            id: "02-35",
          },
          {
            label: "SetTextMargins",
            id: "02-36",
          },
          {
            label: "SetTextTransparency",
            id: "02-37",
          },
          {
            label: "SuperscriptAndSubscript",
            id: "02-38",
          },
        ],
      },
      //03_Slides
      {
        label: "03_Slides",
        children: [
          {
            label: "Layout",
            id: "03-1",
            children: [
              {
                label: "AddDifferentLayoutSlides",
                id: "03-1-1",
              },
              {
                label: "ChangeSlideLayout",
                id: "03-1-2",
              },
              {
                label: "GetSlideLayoutName",
                id: "03-1-3",
              },
              {
                label: "SetSlideLayout",
                id: "03-1-4",
              },
            ]
          },
          {
            label: "Transitions",
            id: "03-2",
            children: [
              {
                label: "BetterSlideTransitions",
                id: "03-2-1",
              },
              {
                label: "SetAdvanceAfterTime",
                id: "03-2-2",
              },
              {
                label: "SetTransitionEffects",
                id: "03-2-3",
              },
              {
                label: "SetTransitions",
                id: "03-2-4",
              },
            ]
          },
          {
            label: "AddImageInMaster",
            id: "03-3",
          },
          {
            label: "AddSlideUsingMasterLayout",
            id: "03-4",
          },
          {
            label: "AppendSlideWithMasterLayout",
            id: "03-5",
          },
          {
            label: "ApplySlideMaster",
            id: "03-6",
          },
          {
            label: "ChangeSlidePosition",
            id: "03-7",
          },
          {
            label: "ClonePPTAtEndOfAnother",
            id: "03-8",
          },
          {
            label: "ClonePPTMasterToAnother",
            id: "03-9",
          },
          {
            label: "CloneSlideAtTheEnd",
            id: "03-10",
          },
          {
            label: "CloneSlideToAnotherPPT",
            id: "03-11",
          },
          {
            label: "CloneSlideWithinAPPT",
            id: "03-12",
          },
          {
            label: "CreateSlide",
            id: "03-13",
          },
          {
            label: "CreateSlideMasterAndApply",
            id: "03-14",
          },
          {
            label: "DetectUsedThemes",
            id: "03-15",
          },
          {
            label: "GetSlideByIndexOrID",
            id: "03-16",
          },
          {
            label: "GetSlideText",
            id: "03-17",
          },
          {
            label: "HideSlide",
            id: "03-18",
          },
          {
            label: "MergeSelectedSlides",
            id: "03-19",
          },
          {
            label: "RemoveSlide",
            id: "03-20",
          },
          {
            label: "RemoveUnusedLayoutMaster",
            id: "03-21",
          },
          {
            label: "SetStartingNumberForSlides",
            id: "03-22",
          },
          {
            label: "SlideTitle",
            id: "03-23",
          },
        ],
      },
      // 04_Shapes
      {
        label: "04_Shapes",
        children: [
          {
            label: "AddLineToSlide",
            id: "4-1",
          },
          {
            label: "AddLineWithArrow",
            id: "4-2",
          },
          {
            label: "AddLineWithTwoPoints",
            id: "4-3",
          },
          {
            label: "AddRoundCornerRectangle",
            id: "4-4",
          },
          {
            label: "AddShapes",
            id: "4-5",
          },
          {
            label: "ArrangeShapes",
            id: "4-6",
          },
          {
            label: "Background",
            id: "4-7",
          },
          {
            label: "CopyShapesBetweenSlides",
            id: "4-8",
          },
          {
            label: "FillShapeWithGradient",
            id: "4-9",
          },
          {
            label: "FillShapeWithPattern",
            id: "4-10",
          },
          {
            label: "FillShapeWithPicture",
            id: "4-11",
          },
          {
            label: "FillShapeWithSolidColor",
            id: "4-12",
          },
          {
            label: "FindShapeByAltText",
            id: "4-13",
          },
          {
            label: "GetAllTitles",
            id: "4-14",
          },
          {
            label: "GetShapeGroupAltText",
            id: "4-15",
          },
          {
            label: "GetShapesByPlaceholder",
            id: "4-16",
          },
          {
            label: "GroupShapes",
            id: "4-17",
          },
          {
            label: "HideShape",
            id: "4-18",
          },
          {
            label: "IsTextBox",
            id: "4-19",
          },
          {
            label: "OperatePlaceholders",
            id: "4-20",
          },
          {
            label: "PreventOrAllowChangingShape",
            id: "4-21",
          },
          {
            label: "RemoveShape",
            id: "4-22",
          },
          {
            label: "ReorderOverlappingShapes",
            id: "4-23",
          },
          {
            label: "ResetPositionOfPlaceholder",
            id: "4-24",
          },
          {
            label: "ResetShapeSizeAndPosition",
            id: "4-25",
          },
          {
            label: "RotateShape",
            id: "4-26",
          },
          {
            label: "Set3DEffectForShape",
            id: "4-27",
          },
          {
            label: "SetAndGetAlternativeText",
            id: "4-28",
          },
          {
            label: "SetEllipseFormat",
            id: "4-29",
          },
          {
            label: "SetFormatForLines",
            id: "4-30",
          },
          {
            label: "SetLineJoinStyles",
            id: "4-31",
          },
          {
            label: "SetOutlineAndEffectForShape",
            id: "4-32",
          },
          {
            label: "SetRadiusForRoundedRectangle",
            id: "4-33",
          },
          {
            label: "SetRadiusOfRoundedRectangle",
            id: "4-34",
          },
          {
            label: "SetRectangleFormat",
            id: "4-35",
          },
          {
            label: "SetShadowEffectForShape",
            id: "4-36",
          },
          {
            label: "ShapeToImage",
            id: "4-37",
          },
          {
            label: "UngroupShapes",
            id: "4-38",
          },
          {
            label: "Animations",
            children: [
              {
                label: "AddExitAnimationForShape",
                id: "4-39-1",
              },
              {
                label: "Animations",
                id: "4-39-2",
              },
              {
                label: "ApplyAnimationOnChart",
                id: "4-39-3",
              },
              {
                label: "ApplyAnimationOnShape",
                id: "4-39-4",
              },
              {
                label: "ApplyAnimationOnText",
                id: "4-39-5",
              },
              {
                label: "CustomPathAnimation",
                id: "4-39-6",
              },
              {
                label: "DurationAndDelayTime",
                id: "4-39-7",
              },
              {
                label: "GetAnimationEffectInfo",
                id: "4-39-8",
              },
              {
                label: "GetAnimationsMotionPath",
                id: "4-39-9",
              },
              {
                label: "SetAnimationForAnimateText",
                id: "4-39-10",
              },
              {
                label: "SetAnimationRepeatType",
                id: "4-39-11",
              },
            ],
          },
        ],
      },
      // 05_DocumentOperation
      {
        label: "05_DocumentOperation",
        children: [
          {
            label: "AddSection",
            id: "5-1",
          },
          {
            label: "AddSlidetoSection",
            id: "5-2",
          },
          {
            label: "DeleteSection",
            id: "5-3",
          },
          {
            label: "GetSectionIndex",
            id: "5-4",
          },
          {
            label: "LoadFromStream",
            id: "5-5",
          },
          {
            label: "LoopPresentPPT",
            id: "5-6",
          },
          {
            label: "PageSetup",
            id: "5-7",
          },
          {
            label: "SaveToStream",
            id: "5-8",
          },
          {
            label: "SetShowTypeAsKiosk",
            id: "5-9",
          },
          {
            label: "SplitPPT",
            id: "5-10",
          },
        ],
      },
      // 06_Properties
      {
        label: "06_Properties",
        children: [
          {
            label: "GetBuiltinProperties",
            id: "6-1",
          },
          {
            label: "MarkAsFinal",
            id: "6-2",
          },
          {
            label: "Properties",
            id: "6-3",
          },
          {
            label: "SetPropertiesForTemplate",
            id: "6-4",
          },
        ],
      },
      // 07_Security
      {
        label: "07_Security",
        children: [
          {
            label: "CheckPasswordProtection",
            id: "7-1",
          },
          {
            label: "Encrypt",
            id: "7-2",
          },
          {
            label: "ModifyPasswordOfEncryptedPPT",
            id: "7-3",
          },
          {
            label: "OpenEncryptedPPT",
            id: "7-4",
          },
          {
            label: "RemoveAllDigitalSignatures",
            id: "7-5",
          },
          {
            label: "RemoveEncryption",
            id: "7-6",
          },
          {
            label: "SetDocumentReadOnly",
            id: "7-7",
          },
        ],
      },
      // 08_Background
      {
        label: "08_Background",
        children: [
          {
            label: "SetBackground",
            id: "8-1",
          },
          {
            label: "SetGradientBackground",
            id: "8-2",
          },
          {
            label: "SetMasterBackground",
            id: "8-3",
          },
        ],
      },
      // 09_Charts
      {
        label: "09_Charts",
        children: [
          {
            label: "AddAndFormatErrorBars",
            id: "9-1",
          },
          {
            label: "AddCustomErrorBars",
            id: "9-2",
          },
          {
            label: "AddSecondaryValueAxis",
            id: "9-3",
          },
          {
            label: "AddShadowEffectForDataLabel",
            id: "9-4",
          },
          {
            label: "AddTrendLineForChartSeries",
            id: "9-5",
          },
          {
            label: "AutoVaryColorForPieChart",
            id: "9-6",
          },
          {
            label: "ChangeColorForLegend",
            id: "9-7",
          },
          {
            label: "ChangeFontForChartDataTable",
            id: "9-8",
          },
          {
            label: "ChangeFontSizeForLegend",
            id: "9-9",
          },
          {
            label: "ChangeSeriesName",
            id: "9-10",
          },
          {
            label: "ChangesForTrendLineEquation",
            id: "9-11",
          },
          {
            label: "ChangeTextFontInChart",
            id: "9-12",
          },
          {
            label: "ChartAxis",
            id: "9-13",
          },
          {
            label: "CopyChartBetweenPptFiles",
            id: "9-14",
          },
          {
            label: "CopyChartWithinOnePPT",
            id: "9-15",
          },
          {
            label: "Create100PercentStackedBarChart",
            id: "9-16",
          },
          {
            label: "CreateBoxAndWhiskerChart",
            id: "9-17",
          },
          {
            label: "CreateBubbleChart",
            id: "9-18",
          },
          {
            label: "CreateClusteredColumnChart",
            id: "9-19",
          },
          {
            label: "CreateCombinationChart",
            id: "9-20",
          },
          {
            label: "CreateCylinder3DClusteredChart",
            id: "9-21",
          },
          {
            label: "CreateDoughnutChart",
            id: "9-22",
          },
          {
            label: "CreateFunnelChart",
            id: "9-23",
          },
          {
            label: "CreateHistogramChart",
            id: "9-24",
          },
          {
            label: "CreateLineMarkersChart",
            id: "9-25",
          },
          {
            label: "CreateMapChart",
            id: "9-26",
          },
          {
            label: "CreateParetoChart",
            id: "9-27",
          },
          {
            label: "CreatePieChart",
            id: "9-28",
          },
          {
            label: "CreateScatterChart",
            id: "9-29",
          },
          {
            label: "CreateSunBurstChart",
            id: "9-30",
          },
          {
            label: "CreateTreeMapChart",
            id: "9-31",
          },
          {
            label: "CreateWaterFallChart",
            id: "9-32",
          },
          {
            label: "DeleteChartLegendEntries",
            id: "9-33",
          },
          {
            label: "DetectIsSwitchRowAndColumn",
            id: "9-34",
          },
          {
            label: "DoughnutChartHole",
            id: "9-35",
          },
          {
            label: "EditChartData",
            id: "9-36",
          },
          {
            label: "ExplodePieChart",
            id: "9-37",
          },
          {
            label: "FillPictureInChartMarker",
            id: "9-38",
          },
          {
            label: "FormatChartDataLabels",
            id: "9-39",
          },
          {
            label: "GetValuesAndUnitFromAxis",
            id: "9-40",
          },
          {
            label: "GroupTwoLevelAxisLabels",
            id: "9-41",
          },
          {
            label: "HideAxisAndGridLine",
            id: "9-42",
          },
          {
            label: "HideOrShowASeriesOfChart",
            id: "9-43",
          },
          {
            label: "InvertIfNegative",
            id: "9-44",
          },
          {
            label: "ModifyChartCategoryAxis",
            id: "9-45",
          },
          {
            label: "MultipleCategoryChart",
            id: "9-46",
          },
          {
            label: "ProtectChart",
            id: "9-47",
          },
          {
            label: "RemoveChartFromPptSlide",
            id: "9-48",
          },
          {
            label: "RemoveTickMarksOfAxis",
            id: "9-49",
          },
          {
            label: "SaveChartAsImage",
            id: "9-50",
          },
          {
            label: "ScaleBubbleChartSize",
            id: "9-51",
          },
          {
            label: "SetAxisPosition",
            id: "9-52",
          },
          {
            label: "SetAxisType",
            id: "9-53",
          },
          {
            label: "SetChartBorderstyle",
            id: "9-54",
          },
          {
            label: "SetChartDataLabelRange",
            id: "9-55",
          },
          {
            label: "SetChartDataNumberFormat",
            id: "9-56",
          },
          {
            label: "SetColorAndNameForTrendline",
            id: "9-57",
          },
          {
            label: "SetDatalabelPosition",
            id: "9-58",
          },
          {
            label: "SetDatapointColorInChart",
            id: "9-59",
          },
          {
            label: "SetDisplayUnit",
            id: "9-60",
          },
          {
            label: "SetDistanceFromAxis",
            id: "9-61",
          },
          {
            label: "SetGapWidth",
            id: "9-62",
          },
          {
            label: "SetLegendOptions",
            id: "9-63",
          },
          {
            label: "SetNumberFormatForAxis",
            id: "9-64",
          },
          {
            label: "SetPercentageForLabels",
            id: "9-65",
          },
          {
            label: "SetPositionOfChartDataLabels",
            id: "9-66",
          },
          {
            label: "SetRotationForChartTitle",
            id: "9-67",
          },
          {
            label: "SetRotationForDataLabel",
            id: "9-68",
          },
          {
            label: "SetRotationForValueAxisText",
            id: "9-69",
          },
          {
            label: "SetSeriesLineColor",
            id: "9-70",
          },
          {
            label: "SetSeriesOverlap",
            id: "9-71",
          },
          {
            label: "SetSizeAndStyleForMarker",
            id: "9-72",
          },
          {
            label: "SetSizeForChartPlotArea",
            id: "9-73",
          },
          {
            label: "SetTextFontForChartTitle",
            id: "9-74",
          },
          {
            label: "SetTextFontForLegendAndAxis",
            id: "9-75",
          },
          {
            label: "SetTickMarkLabelsOnCategoryAxis",
            id: "9-76",
          },
          {
            label: "SetTickMarksInterval",
            id: "9-77",
          },
          {
            label: "ShowChartLabels",
            id: "9-78",
          },
          {
            label: "VaryColorOfSameSerieDataMarker",
            id: "9-79",
          },
        ],
      },
      //10_Conversion
      {
        label: "10_Conversion",
        children: [
          {
            label: "ConvertODPtoPDF",
            id: "10-1",
          },
          {
            label: "ConvertPdfWithDefaultFont",
            id: "10-2",
          },
          {
            label: "ConvertPPSToPPTX",
            id: "10-3",
          },
          {
            label: "ConvertPPTToOFD",
            id: "10-4",
          },
          {
            label: "ConvertUnhideSlidesToPdf",
            id: "10-5",
          },
          {
            label: "CustomSizeToTIFF",
            id: "10-6",
          },
          {
            label: "IndividualSlideToHtml",
            id: "10-7",
          },
          {
            label: "LoadSaveDPSAndDPT",
            id: "10-8",
          },
          {
            label: "OdpToPdf",
            id: "10-9",
          },
          {
            label: "OneSlideToSVG",
            id: "10-10",
          },
          {
            label: "SlideToSVG",
            id: "10-11",
          },
          {
            label: "SpecificSlideToPDF",
            id: "10-12",
          },
          {
            label: "ToHTML",
            id: "10-13",
          },
          {
            label: "ToImage",
            id: "10-14",
          },
          {
            label: "ToMarkDown",
            id: "10-15",
          },
          {
            label: "ToPDF",
            id: "10-16",
          },
          {
            label: "ToPdfWithSpecificPageSize",
            id: "10-17",
          },
          {
            label: "ToPPTX",
            id: "10-18",
          },
          {
            label: "ToSpecificSizeImage",
            id: "10-19",
          },
          {
            label: "ToSVG",
            id: "10-20",
          },
          {
            label: "ToXPS",
            id: "10-21",
          },
        ],
      },
      //11_Images
      {
        label: "11_Images",
        children: [
          {
            label: "ChangeImageSize",
            id: "11-1",
          },
          {
            label: "CropImage",
            id: "11-2",
          },
          {
            label: "ExtractImage",
            id: "11-3",
          },
          {
            label: "ExtractImageFromSpecificSlide",
            id: "11-4",
          },
          {
            label: "InsertEMFInPPT",
            id: "11-5",
          },
          {
            label: "InsertImage",
            id: "11-6",
          },
          {
            label: "RemoveImages",
            id: "11-7",
          },
          {
            label: "SetImageFrameFormat",
            id: "11-8",
          },
          {
            label: "SetImageTransparency",
            id: "11-9",
          },
          {
            label: "UpdateImage",
            id: "11-10",
          },
        ],
      },


      //12_Tables
      {
        label: "12_Tables",
        children: [
          {
            label: "AddImageInTableCell",
            id: "12-1",
          },
          {
            label: "AddRowToTable",
            id: "12-2",
          },
          {
            label: "AdjustColumnByTextWidth",
            id: "12-3",
          },
          {
            label: "CloneRowAndColumn",
            id: "12-4",
          },
          {
            label: "CreateTable",
            id: "12-5",
          },
          {
            label: "EditTableDataAndStyle",
            id: "12-6",
          },
          {
            label: "FillAllCellsWithColor",
            id: "12-7",
          },
          {
            label: "FillParticularRowWithColor",
            id: "12-8",
          },
          {
            label: "GetBorderColorOfCell",
            id: "12-9",
          },
          {
            label: "IdentifyMergedCells",
            id: "12-10",
          },
          {
            label: "LockAspectRatio",
            id: "12-11",
          },

          {
            label: "LockAspectRatio",
            id: "12-11",
          },
          {
            label: "MergeTableCell",
            id: "12-12",
          },
          {
            label: "RemoveRowsAndColumns",
            id: "12-13",
          },
          {
            label: "RemoveTableBorderStyle",
            id: "12-14",
          },
          {
            label: "RemoveTableFromPptSlide",
            id: "12-15",
          },
          {
            label: "SetAlignmentInTable",
            id: "12-16",
          },
          {
            label: "SetBordersForExistingTable",
            id: "12-17",
          },
          {
            label: "SetBordersForNewlyTables",
            id: "12-18",
          },
          {
            label: "SetFirstRowAsHeader",
            id: "12-19",
          },
          {
            label: "SetRowHeightColumnWidth",
            id: "12-20",
          },
          {
            label: "SetTableBorderStyle",
            id: "12-21",
          },
          {
            label: "SetTableStyle",
            id: "12-22",
          },
          {
            label: "SetTextFormat",
            id: "12-23",
          },
          {
            label: "SplitSpecificTableCell",
            id: "12-24",
          },
          {
            label: "TraverseThroughCells",
            id: "12-25",
          },
        ],
      },
      //13_Hyperlink
      {
        label: "13_Hyperlink",
        children: [
          {
            label: "AddHyperlinkToImage",
            id: "13-1",
          },
          {
            label: "AddHyperlinkToSmartArtNode",
            id: "13-2",
          },
          {
            label: "AddHyperlinkToText",
            id: "13-3",
          },
          {
            label: "ChangeHyperlinkColor",
            id: "13-4",
          },
          {
            label: "GetLinkedSlide",
            id: "13-5",
          },
          {
            label: "HyperlinkOutlineStyle",
            id: "13-6",
          },
          {
            label: "Hyperlinks",
            id: "13-7",
          },
          {
            label: "LinkToASpecificSlide",
            id: "13-8",
          },
          {
            label: "LinkToLastViewedSlide",
            id: "13-9",
          },
          {
            label: "ModifyHyperlink",
            id: "13-10",
          },
          {
            label: "RemoveHyperlink",
            id: "13-11",
          },
        ],
      },
      //14_AudioAndVideo
      {
        label: "14_AudioAndVideo",
        children: [
          {
            label: "ExtractAudio",
            id: "14-1",
          },
          {
            label: "ExtractVideo",
            id: "14-2",
          },
          {
            label: "HideAudioDuringShow",
            id: "14-3",
          },
          {
            label: "InsertAudio",
            id: "14-4",
          },
          {
            label: "InsertVideo",
            id: "14-5",
          },
          {
            label: "ObtainSoundEffect",
            id: "14-6",
          },
          {
            label: "ReplaceVideo",
            id: "14-7",
          },
          {
            label: "SetPlayModeForVideo",
            id: "14-8",
          },
        ],
      },
      //15_CommentAndNote
      {
        label: "15_CommentAndNote",
        children: [
          {
            label: "AddAndGetSpeakerNotes",
            id: "15-1",
          },
          {
            label: "AddComment",
            id: "15-2",
          },
          {
            label: "AddNote",
            id: "15-3",
          },
          {
            label: "DeleteComment",
            id: "15-4",
          },
          {
            label: "ExtractComments",
            id: "15-5",
          },
          {
            label: "GetSlideComments",
            id: "15-6",
          },
          {
            label: "PptToSvgRetainNotes",
            id: "15-7",
          },
          {
            label: "RemoveNoteAtSpecificSlide",
            id: "15-8",
          },
          {
            label: "RemoveSpeakerNotes",
            id: "15-9",
          },
          {
            label: "ReplyToComment",
            id: "15-10",
          },
        ],
      },
      //16_HeaderAndFooter
      {
        label: "16_HeaderAndFooter",
        children: [
          {
            label: "HeaderAndFooter",
            id: "16-1",
          },
          {
            label: "ManageNoteMasterHeaderFooter",
            id: "16-2",
          },
        ],
      },
      //17_SmartArt
      {
        label: "17_SmartArt",
        children: [
          {
            label: "AccessChildNodes",
            id: "17-1",
          },
          {
            label: "AccessSmartArt",
            id: "17-2",
          },
          {
            label: "AccessSmartArtLayout",
            id: "17-3",
          },
          {
            label: "AccessSpecificChildNode",
            id: "17-4",
          },
          {
            label: "AddNodeByPosition",
            id: "17-5",
          },
          {
            label: "AddSmartArtNode",
            id: "17-6",
          },
          {
            label: "AssistantNode",
            id: "17-7",
          },
          {
            label: "ChangeNodeText",
            id: "17-8",
          },
          {
            label: "ChangeSmartArtColorStyle",
            id: "17-9",
          },
          {
            label: "ChangeSmartArtShapeStyle",
            id: "17-10",
          },
          {
            label: "CreateSmartArtShape",
            id: "17-11",
          },
          {
            label: "ExtractTextFromSmartArt",
            id: "17-12",
          },
          {
            label: "RemoveNode",
            id: "17-13",
          },
          {
            label: "SetSmartArtLinklineOutline",
            id: "17-14",
          },
          {
            label: "SetSmartArtNodeOutline",
            id: "17-15",
          },
        ],
      },
      //18_Watermark
      {
        label: "18_Watermark",
        children: [
          {
            label: "AddImageWatermark",
            id: "18-1",
          },
          {
            label: "AddWatermark",
            id: "18-2",
          },
          {
            label: "RemoveTextOrImageWatermark",
            id: "18-3",
          },
        ],
      },
      //19_OLE
      {
        label: "19_OLE",
        children: [
          {
            label: "EmbedExcelAsOLE",
            id: "19-1",
          },
          {
            label: "EmbedZipIntoPPT",
            id: "19-2",
          },
          {
            label: "ExtractOLEObject",
            id: "19-3",
          },
          {
            label: "ModifyOLEData",
            id: "19-4",
          },
        ],
      },
      //20_VBA
      {
        label: "20_VBA",
        children: [
          {
            label: "RemoveVBAMacros",
            id: "20-1",
          },
        ],
      },




    ],
  },

];

const defaultProps = {
  children: "children",
  label: "label",
};
</script>