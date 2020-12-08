const imageUrl = "https://media.gettyimages.com/vectors/santa-klaus-jump-kick-vector-id499768808?b=1&k=6&m=499768808&s=170x170&h=O1c06OT0PBVrOI8rVmIzWGq_3n8534TXF60f0SFpG8E="; // Santa kicking

async function main(workbook: ExcelScript.Workbook): Promise<void> {
  let sheet = workbook.addWorksheet();
  sheet.activate();

  let image = await ImageInfo.fromUrl(`https://sofetch.glitch.me/${encodeURI(imageUrl)}`);

  await renderImage(sheet, image);
}

const MAX_IMAGE_WIDTH = 240;
const MAX_IMAGE_HEIGHT = 120;
const UNIT_WIDTH = 30;
const UNIT_HEIGHT = 30;
const CELL_WIDTH = 4;
const CELL_HEIGHT = 4;
const OFFSET_X = 0;
const OFFSET_Y = 0;
const BREATHING_MILLISECONDS = 3000;

async function renderImage(sheet: ExcelScript.Worksheet, image: ImageInfo): Promise<void> {
  console.log(`${image.originalWidth}, ${image.originalHeight}, ${image.resizedWidth}, ${image.resizedHeight}`);

  let address = `${columnToCanonical(OFFSET_X)}${1 + OFFSET_Y}:${columnToCanonical(image.resizedWidth + OFFSET_Y)}${image.resizedHeight + OFFSET_Y}`
  let format = sheet.getRange(address).getFormat();
  format.setColumnWidth(CELL_WIDTH);
  format.setRowHeight(CELL_HEIGHT);

  const rowPasses = Math.ceil(image.resizedHeight / UNIT_HEIGHT);
  const columnPasses = Math.ceil(image.resizedWidth / UNIT_WIDTH);
  let currentColor = "";
  let currentRow = 0;
  let currentColumn = 0;

  try {
    for (let y = 0; y < rowPasses; y++) {
      for (let x = 0; x < columnPasses; x++) {

        console.log(`Rendering segment ${y}, ${x}...`);

        for (let row = y * UNIT_HEIGHT; row < Math.min((y + 1) * UNIT_HEIGHT, image.resizedHeight); row++) {
          for (let column = x * UNIT_WIDTH; column < Math.min((x + 1) * UNIT_WIDTH, image.resizedWidth); column++) {
            const hex =
              decimalToHex(image.pixels[row * image.resizedWidth * 4 + column * 4], 2) +
              decimalToHex(image.pixels[row * image.resizedWidth * 4 + column * 4 + 1], 2) +
              decimalToHex(image.pixels[row * image.resizedWidth * 4 + column * 4 + 2], 2);

            currentColor = hex;
            currentRow = row;
            currentColumn = column;
            let cell = sheet.getCell(row + OFFSET_Y, column + OFFSET_X);
            cell.getFormat().getFill().setColor(hex);
          }
        }

        if (y < rowPasses - 1 || x < columnPasses - 1) {
          console.log("Breathing...");
          await sleep(BREATHING_MILLISECONDS);
        }
      }
    }
  } catch (ex) {
    console.log("Failed to render!", ex, `(${currentColumn}, ${currentRow}) - ${currentColor}`);
  }
}

class ImageInfo {
  pixels: Uint8ClampedArray | null;
  originalWidth: number;
  originalHeight: number;
  resizedWidth: number;
  resizedHeight: number;

  static async fromUrl(imageUrl: string): Promise<ImageInfo> {
    let fetchResult = await fetch(`https://sofetch.glitch.me/${encodeURI(imageUrl)}`);
    let blob = await fetchResult.blob();
    let imageBitmap = await globalThis.createImageBitmap(blob);
    return new ImageInfo(imageBitmap);
  }

  private constructor(imageBitmap: ImageBitmap) {
    this.originalWidth = imageBitmap.width;
    this.originalHeight = imageBitmap.height;
    this.adjustDimension();
    this.pixels = ImageInfo.getPixels(imageBitmap, this.resizedWidth, this.resizedHeight);
  }

  private adjustDimension() {
    this.resizedWidth = this.originalWidth;
    this.resizedHeight = this.originalHeight;

    if (this.resizedWidth > MAX_IMAGE_WIDTH) {
      this.resizedWidth = MAX_IMAGE_WIDTH;
      this.resizedHeight = Math.ceil(this.resizedWidth * this.originalHeight / this.originalWidth);
    }

    if (this.resizedHeight > MAX_IMAGE_HEIGHT) {
      this.resizedHeight = MAX_IMAGE_HEIGHT;
      this.resizedWidth = Math.ceil(this.resizedHeight * this.originalWidth / this.originalHeight);
    }
  }

  private static getPixels(imageBitmap: ImageBitmap, targetWidth: number, targetHeight: number): Uint8ClampedArray {
    let canvas = new globalThis.OffscreenCanvas(targetWidth, targetHeight);
    let context2d = canvas.getContext("2d");
    context2d.drawImage(imageBitmap, 0, 0, imageBitmap.width, imageBitmap.height, 0, 0, targetWidth, targetHeight);
    return context2d.getImageData(0, 0, targetWidth, targetHeight).data;
  }
}

function sleep(milliseconds: number) {
  return new Promise(resolve => setTimeout(resolve, milliseconds));
}

function decimalToHex(decimal: number, padding: number = 2): string {
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