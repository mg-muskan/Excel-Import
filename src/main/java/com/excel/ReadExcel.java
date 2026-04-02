package com.excel;

import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import javax.imageio.IIOImage;
import javax.imageio.ImageIO;
import javax.imageio.ImageWriteParam;
import javax.imageio.ImageWriter;
import javax.imageio.stream.ImageOutputStream;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ReadExcel {

	private static final String S3_IMAGE_URL = "https://s3.ap-south-1.amazonaws.com/_StoreName_/web/stores/%s/media/products/%s";
	private static final String DIR_CANONICALPATH = "_StoreName_/web/stores/%s/media/products";
	private static final String IMAGE_BASE_CANONICALPATH = "_StoreName_/web/stores/%s/media/products/%s";
	private static final String CATALOGUE = "productCatalogue";
	private static final String VERSION = "Staged";
	private static final String STORE_CODE = "green-hermitage";
	private static final String DATAFILE = "C:\\Users\\abc\\Downloads\\green-hermitage.xlsx";
	private static final String EXPORTFILE = "C:\\Users\\abc\\Downloads\\green-hermitage-import.xlsx";
	private static final String OUTPUT_FILEPATH = "C:\\Users\\abc\\Downloads\\green-hermitage\\";
	
	private static final int NEW_WIDTH = 1000;
	private static final float QUALITY = 0.4f;

	public static void main(String[] args) {
		try (InputStream inputstream = FileUtils.openInputStream(new File(DATAFILE))) {

			Workbook inputWorkbook = new XSSFWorkbook(inputstream);
			Sheet inputSheet = inputWorkbook.getSheetAt(0);

			Workbook outputWorkbook = new XSSFWorkbook(); // Use XSSFWorkbook for .xlsx format
			Sheet outputSheet = outputWorkbook.createSheet("data");
			int rowCount = 0;

			outputSheet.createRow(rowCount++).createCell(0)
					.setCellValue("&OrderableUnit=com.xhopfront.entities.OrderableUnit");
			outputSheet.createRow(rowCount++).createCell(0)
					.setCellValue("&Image=com.cabin4j.suite.entity.platform.Image");
			outputSheet.createRow(rowCount++).createCell(0)
					.setCellValue("&GroupingType=com.xhopfront.entities.GroupingType");
			outputSheet.createRow(rowCount++).createCell(0).setCellValue("&Grouping=com.xhopfront.entities.Grouping");
			outputSheet.createRow(rowCount++).createCell(0).setCellValue("&SKU=com.xhopfront.entities.SKU");
			outputSheet.createRow(rowCount++).createCell(0).setCellValue("&PSU=com.xhopfront.entities.PSU");
			outputSheet.createRow(rowCount++).createCell(0).setCellValue("");
			outputSheet.createRow(rowCount++).createCell(0).setCellValue("");

			boolean firstSkipped = false;
			for (Row row : inputSheet) {
				if (!firstSkipped) {
					firstSkipped = true;
					continue;
				}
				String code = null;
				String name = null;
				String grouping1 = null;
				String grouping2 = null;
				String hsncode = null;
				String summary = null;
				String description = null;
				String seoTitle = null;
				String seoDescription = null;
				String seoKeywords = null;
				String image1 = null;
				String image2 = null;
				String image3 = null;
				String price = null;
				String discountprice = null;
				String availableStock = null;
				String gst = null;
				String cancelable = null;
				String returnable = null;
				String returnPeriod = null;
				for (Cell cell : row) {
					
//					System.out.println("##########" + CellReference.convertNumToColString(cell.getColumnIndex()));
//					System.out.println(getCellValueAsString(cell));

					if ("A".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						String decimalString = getCellValueAsString(cell);
						if (decimalString.contains(".")) {
							code = decimalString.substring(0, decimalString.indexOf('.'));
						} else {
							code = decimalString;
						}
					}
					if ("B".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						name = getCellValueAsString(cell);
					}
					if ("C".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						grouping1 = getCellValueAsString(cell).toLowerCase().replaceAll(" ", "-");
					}
					if ("D".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						grouping2 = getCellValueAsString(cell);
					}
					if ("E".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						hsncode = getCellValueAsString(cell);
					}
					if ("F".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						summary = getCellValueAsString(cell);
					}
					if ("G".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						description = getCellValueAsString(cell);
					}
					if ("H".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						seoTitle = getCellValueAsString(cell);
					}
					if ("I".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						seoDescription = getCellValueAsString(cell);
					}
					if ("J".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						seoKeywords = getCellValueAsString(cell);
					}
					if ("K".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						image1 = getCellValueAsString(cell);
					}
					if ("L".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						image2 = getCellValueAsString(cell);
					}
					if ("M".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						image3 = getCellValueAsString(cell);
					}
					if ("N".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						price = getCellValueAsString(cell);
					}
					if ("O".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						discountprice = getCellValueAsString(cell);
					}
					if ("P".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						availableStock = getCellValueAsString(cell);
					}
					if ("Q".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						gst = getCellValueAsString(cell);
					}
					if ("R".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						cancelable = getCellValueAsString(cell);
					}
					if ("S".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						returnable = getCellValueAsString(cell);
					}
					if ("T".equals(CellReference.convertNumToColString(cell.getColumnIndex()))) {
						String decimalString = getCellValueAsString(cell);
						if (decimalString.contains(".")) {
							returnPeriod = decimalString.substring(0, decimalString.indexOf('.'));
						} else {
							returnPeriod = decimalString;
						}
					}
				}
				Row row1 = outputSheet.createRow(rowCount++);
				row1.createCell(0).setCellValue("UPSERT &Image");
				row1.createCell(1).setCellValue("canonicalPath(unique=true)");
				row1.createCell(2).setCellValue("directory(attribute=canonicalPath)");
				row1.createCell(3).setCellValue("url");
				row1.createCell(4).setCellValue("fileName");
				row1.createCell(5).setCellValue("contentType");
				row1.createCell(6).setCellValue("publicAccess");
				row1.createCell(7).setCellValue("size");
				List<String> images = new ArrayList<>();
				
				if (StringUtils.isNotBlank(image1)) {
					String filename = code + "_1.jpg"; 
					if(downloadAndOptimizeImage(image1, OUTPUT_FILEPATH + filename)) {
						images.add(filename);
					}
				}
				if (StringUtils.isNotBlank(image2)) {
					String filename = code + "_2.jpg"; 
					if(downloadAndOptimizeImage(image2, OUTPUT_FILEPATH + filename)) {
						images.add(filename);
					}
				}
				if (StringUtils.isNotBlank(image3)) {
					String filename = code + "_3.jpg"; 
					if(downloadAndOptimizeImage(image3, OUTPUT_FILEPATH + filename)) {
						images.add(filename);
					}
				}	
				for (String img : images) {
					Row row2 = outputSheet.createRow(rowCount++);
					row2.createCell(0).setCellValue("");
					row2.createCell(1).setCellValue(String.format(IMAGE_BASE_CANONICALPATH, STORE_CODE, img));
					row2.createCell(2).setCellValue(String.format(DIR_CANONICALPATH, STORE_CODE));
					row2.createCell(3).setCellValue(String.format(S3_IMAGE_URL, STORE_CODE, img));
					row2.createCell(4).setCellValue(img);
					row2.createCell(5).setCellValue("image/jpeg");
					row2.createCell(6).setCellValue("true");
					row2.createCell(7).setCellValue("100000");
				}

				outputSheet.createRow(rowCount++).createCell(0).setCellValue("");
				
				Row row21 = outputSheet.createRow(rowCount++);
				row21.createCell(0).setCellValue("FETCH &Grouping");
				row21.createCell(1).setCellValue("code(unique=true)");
				row21.createCell(2).setCellValue("catalogue(attribute=code,unique=true)");
				row21.createCell(3).setCellValue("version(unique=true)");
				
				if(StringUtils.isNotBlank(grouping1)) {
					Row row22 = outputSheet.createRow(rowCount++);
					row22.createCell(0).setCellValue(grouping1);
					row22.createCell(1).setCellValue(grouping1);
					row22.createCell(2).setCellValue(CATALOGUE);
					row22.createCell(3).setCellValue(VERSION);
				}
				
				if(StringUtils.isNotBlank(grouping2)) {
					Row row22 = outputSheet.createRow(rowCount++);
					row22.createCell(0).setCellValue(grouping2);
					row22.createCell(1).setCellValue(grouping2);
					row22.createCell(2).setCellValue(CATALOGUE);
					row22.createCell(3).setCellValue(VERSION);
				}
				
				outputSheet.createRow(rowCount++).createCell(0).setCellValue("");
			
				Row row2 = outputSheet.createRow(rowCount++);
				row2.createCell(0).setCellValue("UPSERT &SKU");
				row2.createCell(1).setCellValue("code(unique=true)");
				row2.createCell(2).setCellValue("catalogue(attribute=code,unique=true)");
				row2.createCell(3).setCellValue("version(unique=true)");
				row2.createCell(4).setCellValue("name");
				row2.createCell(5).setCellValue("active");
				row2.createCell(6).setCellValue("searchable");
				row2.createCell(7).setCellValue("store(attribute=code)");
				row2.createCell(8).setCellValue("hsnCode");
				row2.createCell(9).setCellValue("summary");
				row2.createCell(10).setCellValue("description");
				row2.createCell(11).setCellValue("groupings(attribute=*)");
				row2.createCell(12).setCellValue("seoTitle");
				row2.createCell(13).setCellValue("seoDescription");
				row2.createCell(14).setCellValue("seoKeywords");
				row2.createCell(15).setCellValue("thumbnail(attribute=canonicalPath)");
				row2.createCell(16).setCellValue("images(attribute=canonicalPath)");
				row2.createCell(17).setCellValue("gstCategory(attribute=code)");
				row2.createCell(18).setCellValue("cancellable");
				row2.createCell(19).setCellValue("returnable");
				row2.createCell(20).setCellValue("returnPeriod");

				Row row3 = outputSheet.createRow(rowCount++);
				row3.createCell(0).setCellValue("SKU-" + code);
				row3.createCell(1).setCellValue(code);
				row3.createCell(2).setCellValue(CATALOGUE);
				row3.createCell(3).setCellValue(VERSION);
				row3.createCell(4).setCellValue(name);
				row3.createCell(5).setCellValue("true");
				row3.createCell(6).setCellValue("true");
				row3.createCell(7).setCellValue(STORE_CODE);
				row3.createCell(8).setCellValue(hsncode);
				row3.createCell(9).setCellValue(summary);
				row3.createCell(10).setCellValue(description);
				row3.createCell(11).setCellValue(Stream.of(grouping1, grouping2).filter(StringUtils::isNotBlank).collect(Collectors.joining(",")));
				row3.createCell(12).setCellValue(seoTitle);
				row3.createCell(13).setCellValue(seoDescription);
				row3.createCell(14).setCellValue(seoKeywords);
				if(CollectionUtils.isNotEmpty(images)) {
					row3.createCell(15).setCellValue(String.format(IMAGE_BASE_CANONICALPATH, STORE_CODE, images.get(0)));
					String imgs = images.stream().filter(StringUtils::isNotBlank).map(img -> String.format(IMAGE_BASE_CANONICALPATH, STORE_CODE, img)).collect(Collectors.joining(","));
					row3.createCell(16).setCellValue(imgs);
				} else {
					row3.createCell(15).setCellValue("");
					row3.createCell(16).setCellValue("");
				}
				row3.createCell(17).setCellValue(gst);
				row3.createCell(18).setCellValue(StringUtils.isNotBlank(cancelable) ? ("YES".equalsIgnoreCase(cancelable) ? "true" : "false") : "false");
				row3.createCell(19).setCellValue(StringUtils.isNotBlank(returnable) ? ("YES".equalsIgnoreCase(returnable) ? "true" : "false") : "false");
				row3.createCell(20).setCellValue(returnPeriod);

				outputSheet.createRow(rowCount++).createCell(0).setCellValue("");


				// PSU
				Row row4 = outputSheet.createRow(rowCount++);
				row4.createCell(0).setCellValue("UPSERT &PSU");
				row4.createCell(1).setCellValue("product(attribute=*,unique=true)");
				row4.createCell(2).setCellValue("unit(attribute=code,unique=true)");
				row4.createCell(3).setCellValue("currency(attribute=isocode,unique=true)");
				row4.createCell(4).setCellValue("price");
				row4.createCell(5).setCellValue("discountPrice");
				row4.createCell(6).setCellValue("availableStock");
				row4.createCell(7).setCellValue("minimumQuantity");

				Row row5 = outputSheet.createRow(rowCount++);
				row5.createCell(0).setCellValue("");
				row5.createCell(1).setCellValue("SKU-" + code);
				row5.createCell(2).setCellValue("piece");
				row5.createCell(3).setCellValue("INR");
				row5.createCell(4).setCellValue(price);
				row5.createCell(5).setCellValue(discountprice);
				row5.createCell(6).setCellValue(availableStock);
				row5.createCell(7).setCellValue("1");

				outputSheet.createRow(rowCount++).createCell(0).setCellValue("");
				outputSheet.createRow(rowCount++).createCell(0).setCellValue("");
			}

			try (FileOutputStream outputStream = new FileOutputStream(EXPORTFILE)) {
				outputWorkbook.write(outputStream);
				System.out.println("Excel file generated successfully!");
			} catch (IOException e) {
				e.printStackTrace();
			} finally {
				// Close the workbook to release resources
				try {
					outputWorkbook.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			// Close the workbook and input stream
			inputWorkbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static String getCellValueAsString(Cell cell) {
		switch (cell.getCellType()) {
		case STRING:
			return cell.getStringCellValue();
		case NUMERIC:
			return String.valueOf(cell.getNumericCellValue());
		case BOOLEAN:
			return Boolean.toString(cell.getBooleanCellValue());
		case BLANK:
			return StringUtils.EMPTY;
		default:
			return StringUtils.EMPTY;

		}
	}

	public static boolean downloadAndOptimizeImage(String driveUrl, String outputFilePath) {
		String fileId = extractFileId(driveUrl);
		String downloadUrl = "https://drive.google.com/uc?export=download&id=" + fileId;
		String tempFilePath = outputFilePath + ".temp";
		//return true;
		try {
			// Download the image from Google Drive
			downloadFile(downloadUrl, tempFilePath);
			System.out.println("Image downloaded successfully: " + tempFilePath);

			// Optimize the image and save to the output file path
			optimize(tempFilePath, outputFilePath);

			// Delete the original downloaded file
			new File(tempFilePath).delete();
			System.out.println("Temporary file deleted: " + tempFilePath);
			return true;
		} catch (IOException e) {
			System.err.println("Failed to download or optimize image: " + e.getMessage());
			return false;
		}
	}

	private static String extractFileId(String driveUrl) {
		String fileId = "";
		String[] parts = driveUrl.split("/");
		for (int i = 0; i < parts.length; i++) {
			if (parts[i].equals("d")) {
				fileId = parts[i + 1];
				break;
			}
		}
		return fileId;
	}

	private static void downloadFile(String fileUrl, String outputFilePath) throws IOException {
		URL url = new URL(fileUrl);
		HttpURLConnection httpConn = (HttpURLConnection) url.openConnection();
		httpConn.setRequestMethod("GET");

		int responseCode = httpConn.getResponseCode();
		if (responseCode == HttpURLConnection.HTTP_OK) {
			InputStream inputStream = new BufferedInputStream(httpConn.getInputStream());
			FileOutputStream outputStream = new FileOutputStream(outputFilePath);

			byte[] buffer = new byte[4096];
			int bytesRead;
			while ((bytesRead = inputStream.read(buffer)) != -1) {
				outputStream.write(buffer, 0, bytesRead);
			}

			outputStream.close();
			inputStream.close();
		} else {
			throw new IOException("Failed to download file: HTTP response code " + responseCode);
		}

		httpConn.disconnect();
	}

	private static void optimize(String inputFilePath, String outputFilePath) throws IOException {
		// Load the image
		File inputFile = new File(inputFilePath);
		BufferedImage inputImage = ImageIO.read(inputFile);

		// Resize the image
		int scaledWidth = NEW_WIDTH;
		int scaledHeight = (int) (inputImage.getHeight() * (NEW_WIDTH * 1.0 / inputImage.getWidth()));
		BufferedImage outputImage = new BufferedImage(scaledWidth, scaledHeight, inputImage.getType());

		Graphics2D g2d = outputImage.createGraphics();
		g2d.drawImage(inputImage, 0, 0, scaledWidth, scaledHeight, null);
		g2d.dispose();

		// Save the resized image with specified quality
		File outputFile = new File(outputFilePath);
		saveImageWithQuality(outputImage, outputFile, QUALITY);
	}

	private static void saveImageWithQuality(BufferedImage image, File output, float quality) throws IOException {
		// Get a jpeg writer
		Iterator<ImageWriter> writers = ImageIO.getImageWritersByFormatName("jpg");
		if (!writers.hasNext()) {
			throw new IllegalStateException("No writers found");
		}
		ImageWriter writer = writers.next();

		// Create an ImageOutputStream to write to the file
		try (ImageOutputStream ios = ImageIO.createImageOutputStream(output)) {
			writer.setOutput(ios);

			// Set the compression quality
			ImageWriteParam param = writer.getDefaultWriteParam();
			if (param.canWriteCompressed()) {
				param.setCompressionMode(ImageWriteParam.MODE_EXPLICIT);
				param.setCompressionQuality(quality);
			}

			// Write the image
			writer.write(null, new IIOImage(image, null, null), param);
		} finally {
			writer.dispose();
		}
	}
}