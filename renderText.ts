const text = "Office Scripts Rocks!";
const fontSize = 24;
const margin = 4;
const renderRainbowColor = true;
const textColor = "blue";

async function main(workbook: ExcelScript.Workbook): Promise<void> {
  let sheet = workbook.addWorksheet();
  sheet.activate();

  await renderText(sheet, text);
}

const MAX_IMAGE_WIDTH = 120;
const MAX_IMAGE_HEIGHT = 100;
const UNIT_WIDTH = 30;
const UNIT_HEIGHT = 30;
const CELL_WIDTH = 4;
const CELL_HEIGHT = 4;
const OFFSET_X = 0;
const OFFSET_Y = 0;
const BREATHING_MILLISECONDS = 3000;

async function renderText(sheet: ExcelScript.Worksheet, text: string): Promise<void> {

  const image = ImageInfo.fromText(text);

  let address = `${columnToCanonical(0)}${1}:${columnToCanonical(image.width + OFFSET_X - 1)}${image.height + OFFSET_Y}`
  let canvas = sheet.getRange(address);
  let format = canvas.getFormat();
  format.setColumnWidth(CELL_WIDTH);
  format.setRowHeight(CELL_HEIGHT);

  const rowPasses = Math.ceil(image.height / UNIT_HEIGHT);
  const columnPasses = Math.ceil(image.width / UNIT_WIDTH);
  let currentColor = "";
  let currentRow = 0;
  let currentColumn = 0;

  try {
    for (let y = 0; y < rowPasses; y++) {
      for (let x = 0; x < columnPasses; x++) {

        for (let row = y * UNIT_HEIGHT; row < Math.min((y + 1) * UNIT_HEIGHT, image.height); row++) {
          for (let column = x * UNIT_WIDTH; column < Math.min((x + 1) * UNIT_WIDTH, image.width); column++) {
            const red = decimalToHex(image.pixels[row * image.width * 4 + column * 4], 2);
            const green = decimalToHex(image.pixels[row * image.width * 4 + column * 4 + 1], 2);
            const blue = decimalToHex(image.pixels[row * image.width * 4 + column * 4 + 2], 2);
            let hex = `#${red}${green}${blue}`.toLowerCase();

            if (hex === '#ffffff' || hex === '#nannannan') {
              continue;
            }

            currentColor = hex;
            currentRow = row;
            currentColumn = column;
            let cell = sheet.getCell(row + OFFSET_Y, column + OFFSET_X);
            cell.getFormat().getFill().setColor(hex);
          }
        }

        if (y < rowPasses - 1 || x < columnPasses - 1) {
          // console.log("Breathing...");
          await sleep(BREATHING_MILLISECONDS);
        }
      }
    }
  } catch (ex) {
    console.log("Failed to render!", ex, `(${currentColumn}, ${currentRow}) - ${currentColor}`);
  } finally {
    console.log("DONE!");
  }
}

class ImageInfo {
  public static fromText(text: string): ImageInfo {
    let canvas = new globalThis.OffscreenCanvas(2048, 2048);
    let context2d = canvas.getContext("2d");
    context2d.fillStyle = "white";
    context2d.fillRect(0, 0, canvas.width, canvas.height);
    context2d.font = `italic bold ${fontSize}px Times New Roman`;
    context2d.textBaseline = "bottom";
    const metrics = context2d.measureText(text);

    const textWidth = Math.ceil(metrics.width);
    const textHeight = fontSize;

    if (renderRainbowColor) {
      var gradient = context2d.createLinearGradient(0, 0, textWidth, textHeight);
      gradient.addColorStop(0, "rgb(255, 0, 0)");
      gradient.addColorStop(0.5, "rgb(0, 255, 0)");
      gradient.addColorStop(1, "rgb(0, 0, 255)");
      context2d.fillStyle = gradient;
    } else {
      context2d.fillStyle = textColor;
    }

    context2d.fillText(text, margin, textHeight + margin);

    const imageWidth = textWidth + margin * 2;
    const imageHeight = textHeight + margin * 2;

    return new ImageInfo(
      context2d.getImageData(0, 0, imageWidth, imageHeight).data,
      imageWidth, imageHeight
    );
  }

  private constructor(
    public readonly pixels: Uint8ClampedArray | null,
    public readonly width: number,
    public readonly height: number) {
  }
}

function sleep(milliseconds: number) {
  return new Promise(resolve => setTimeout(resolve, milliseconds));
}

function decimalToHex(decimal: number, padding: number = 2): string {
  decimal = Math.max(0, Math.min(255, decimal));
  var hex = Number(decimal).toString(16);
  while (hex.length < padding) {
    hex = "0" + hex;
  }
  return hex;
}

function columnToCanonical(column: number): string {
  let column_part = "";
  let cur: number = column;
  while (cur >= 0) {
    column_part = String.fromCharCode("A".charCodeAt(0) + (cur % 26)) + column_part;
    cur = Math.floor(cur / 26) - 1;
  }
  return column_part;
}
