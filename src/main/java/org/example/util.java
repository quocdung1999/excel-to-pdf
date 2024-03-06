package org.example;

import static org.apache.poi.util.Units.*;
public class util {

    public static double defaultColumnWidthPixels(int width) {
        return width * 1.0 * EMU_PER_CHARACTER / EMU_PER_PIXEL;
    }
    public static double defaultRowHeightPixels(short height) {
        return height * EMU_PER_CHARACTER / 20.0  / EMU_PER_PIXEL;
    }

    public static float pixelsToPoints(float point) {
        return (float) (point * 1.0 * EMU_PER_PIXEL / EMU_PER_POINT);
    }
}
