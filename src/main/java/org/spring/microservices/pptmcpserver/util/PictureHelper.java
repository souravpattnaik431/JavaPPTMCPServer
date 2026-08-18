package org.spring.microservices.pptmcpserver.util;

import org.apache.poi.sl.usermodel.PictureData;

/**
 * Utility helper for resolving image format types based on file extensions.
 */
public final class PictureHelper {

    private PictureHelper() {
        throw new UnsupportedOperationException("PictureHelper is a utility class and cannot be instantiated.");
    }

    /**
     * Determines the Apache POI PictureType from a filename extension using enhanced switch expressions.
     *
     * @param fileName Name or path of the image file.
     * @return Resolved PictureData.PictureType enum.
     */
    public static PictureData.PictureType determinePictureType(String fileName) {
        if (fileName == null || fileName.lastIndexOf('.') == -1) {
            return PictureData.PictureType.PNG;
        }

        String extension = fileName.substring(fileName.lastIndexOf('.') + 1).toLowerCase().trim();

        return switch (extension) {
            case "png" -> PictureData.PictureType.PNG;
            case "jpg", "jpeg" -> PictureData.PictureType.JPEG;
            case "gif" -> PictureData.PictureType.GIF;
            case "bmp" -> PictureData.PictureType.BMP;
            case "svg" -> PictureData.PictureType.SVG;
            case "tif", "tiff" -> PictureData.PictureType.TIFF;
            case "eps" -> PictureData.PictureType.EPS;
            case "wmf" -> PictureData.PictureType.WMF;
            case "emf" -> PictureData.PictureType.EMF;
            default -> PictureData.PictureType.PNG;
        };
    }
}
