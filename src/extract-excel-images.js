#!/usr/bin/env node

import ExcelJS from "exceljs";
import fs from "node:fs";
import fsp from "node:fs/promises";
import path from "node:path";

function columnName(columnIndex) {
    // columnIndex is zero-based: 0 = A, 1 = B, etc.
    let n = columnIndex + 1;
    let result = "";

    while (n > 0) {
        const remainder = (n - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        n = Math.floor((n - 1) / 26);
    }

    return result;
}

function safeDirectoryName(name) {
    return name.replace(/[<>:"/\\|?*]/g, "_").trim();
}

function normaliseExtension(extension) {
    if (!extension) {
        return "bin";
    }

    extension = extension.toLowerCase();

    if (extension === "jpeg") {
        return "jpg";
    }

    if (extension === "tiff") {
        return "tif";
    }

    return extension;
}

function getImageData(media) {
    if (media.buffer) {
        return media.buffer;
    }

    if (media.base64) {
        const base64 = media.base64.replace(
            /^data:.*?;base64,/,
            ""
        );

        return Buffer.from(base64, "base64");
    }

    return null;
}

function uniqueFilename(directory, cell, extension) {
    let filename = path.join(
        directory,
        `${cell}.${extension}`
    );

    if (!fs.existsSync(filename)) {
        return filename;
    }

    let counter = 2;

    while (true) {
        filename = path.join(
            directory,
            `${cell}-${counter}.${extension}`
        );

        if (!fs.existsSync(filename)) {
            return filename;
        }

        counter++;
    }
}

async function extractImages(workbookFile, outputDirectory) {
    const workbook = new ExcelJS.Workbook();

    console.log(`Reading ${workbookFile}...`);

    await workbook.xlsx.readFile(workbookFile);

    await fsp.mkdir(outputDirectory, {
        recursive: true
    });

    let totalImages = 0;

    for (const worksheet of workbook.worksheets) {
        const images = worksheet.getImages();

        console.log(
            `${worksheet.name}: ${images.length} image(s)`
        );

        if (images.length === 0) {
            continue;
        }

        const sheetDirectory = path.join(
            outputDirectory,
            safeDirectoryName(worksheet.name)
        );

        await fsp.mkdir(sheetDirectory, {
            recursive: true
        });

        for (const image of images) {
            const topLeft = image.range?.tl;

            if (!topLeft) {
                console.warn(
                    "  Could not determine image position."
                );
                continue;
            }

            /*
             * nativeCol/nativeRow are zero-based.
             *
             * Fall back to col/row for workbooks where the
             * native coordinates aren't populated.
             */
            const column =
                Number.isInteger(topLeft.nativeCol)
                    ? topLeft.nativeCol
                    : Math.floor(topLeft.col);

            const row =
                Number.isInteger(topLeft.nativeRow)
                    ? topLeft.nativeRow
                    : Math.floor(topLeft.row);

            const cell =
                `${columnName(column)}${row + 1}`;

            /*
             * ExcelJS keeps the actual image data in the
             * workbook media collection.
             */
            const media = workbook.model.media.find(
                item => item.index === image.imageId
            );

            if (!media) {
                console.warn(
                    `  ${cell}: image data not found.`
                );
                continue;
            }

            const imageData = getImageData(media);

            if (!imageData) {
                console.warn(
                    `  ${cell}: no image buffer available.`
                );
                continue;
            }

            const extension = normaliseExtension(
                media.extension ||
                path.extname(media.name || "")
                    .replace(".", "")
            );

            const filename = uniqueFilename(
                sheetDirectory,
                cell,
                extension
            );

            await fsp.writeFile(
                filename,
                imageData
            );

            console.log(
                `  ${cell} -> ${path.relative(
                    outputDirectory,
                    filename
                )}`
            );

            totalImages++;
        }
    }

    console.log();
    console.log(
        `Extracted ${totalImages} image(s).`
    );
    console.log(
        `Output directory: ${path.resolve(outputDirectory)}`
    );

    /*
     * Useful warning for workbooks where ExcelJS sees
     * image files but cannot associate them with cells.
     */
    if (
        totalImages === 0 &&
        workbook.model.media?.length > 0
    ) {
        console.warn();
        console.warn(
            "WARNING: The workbook contains embedded media, " +
            "but ExcelJS could not associate the images with cells."
        );
        console.warn(
            "The workbook may use Excel's newer " +
            "\"Picture in Cell\" format."
        );
    }
}

async function main() {
    const workbookFile = process.argv[2];
    const outputDirectory =
        process.argv[3] || "excel-images";

    if (!workbookFile) {
        console.error(
            "Usage: node extract-excel-images.js " +
            "<workbook.xlsx> [output-directory]"
        );

        process.exit(1);
    }

    try {
        await extractImages(
            workbookFile,
            outputDirectory
        );
    } catch (error) {
        console.error(error);
        process.exit(1);
    }
}

main();