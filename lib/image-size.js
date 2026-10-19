/**
 * Pixel dimensions of an embedded image, read from its own header bytes.
 *
 * The figure rule table carries `minDpi` / `recommendedDpiRange`, and until now
 * nothing could check them: effective DPI is pixel width divided by the width
 * the picture is *placed* at, so both the media bytes and the drawing extent
 * are needed. The extent is already in document.xml (`wp:extent` cx/cy);
 * this module supplies the missing half.
 *
 * Only the formats Word actually embeds are handled, and every parser returns
 * null instead of throwing: a figure whose DPI cannot be computed must not turn
 * into an audit crash, because a crash in one check group silently drops the
 * rest of that group's findings.
 */

const EMU_PER_INCH = 914400;

/** PNG: 8-byte signature, then an IHDR chunk with width/height as uint32 BE. */
function pngSize(buf) {
  if (buf.length < 24) return null;
  if (buf.readUInt32BE(0) !== 0x89504e47) return null;
  return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
}

/** GIF: logical screen descriptor at offset 6, uint16 LE. */
function gifSize(buf) {
  if (buf.length < 10) return null;
  if (buf.toString('latin1', 0, 3) !== 'GIF') return null;
  return { width: buf.readUInt16LE(6), height: buf.readUInt16LE(8) };
}

/** BMP: BITMAPINFOHEADER starts at 14, width/height int32 LE. */
function bmpSize(buf) {
  if (buf.length < 26) return null;
  if (buf.toString('latin1', 0, 2) !== 'BM') return null;
  return { width: Math.abs(buf.readInt32LE(18)), height: Math.abs(buf.readInt32LE(22)) };
}

/**
 * JPEG: walk the marker chain to the first SOFn frame header.
 *
 * Sizes are in the frame segment, but the segment is not at a fixed offset —
 * Exif and ICC blocks sit between SOI and SOF — so the chain has to be walked.
 */
function jpegSize(buf) {
  if (buf.length < 4 || buf[0] !== 0xff || buf[1] !== 0xd8) return null;
  let i = 2;
  while (i + 9 < buf.length) {
    if (buf[i] !== 0xff) {
      i++;
      continue;
    }
    const marker = buf[i + 1];
    // Standalone markers carry no length field.
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      i += 2;
      continue;
    }
    const len = buf.readUInt16BE(i + 2);
    const isSof = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isSof) {
      return { height: buf.readUInt16BE(i + 5), width: buf.readUInt16BE(i + 7) };
    }
    i += 2 + len;
  }
  return null;
}

/** Pixel size of an image buffer, or null when the format is unknown. */
export function imagePixelSize(buf) {
  if (!buf || buf.length < 10) return null;
  return pngSize(buf) ?? jpegSize(buf) ?? gifSize(buf) ?? bmpSize(buf);
}

/**
 * Effective print resolution of a placed image.
 *
 * `cx`/`cy` are the `wp:extent` values in EMU, i.e. the size on the page, not
 * the size in pixels. A picture shown at 10cm wide with 500 pixels therefore
 * lands at 127 dpi regardless of anything the file claims about itself.
 */
export function effectiveDpi(pixel, extent) {
  if (!pixel || !extent) return null;
  const widthInches = extent.cx / EMU_PER_INCH;
  const heightInches = extent.cy / EMU_PER_INCH;
  if (!(widthInches > 0) || !(heightInches > 0)) return null;
  return {
    dpiX: pixel.width / widthInches,
    dpiY: pixel.height / heightInches,
    widthPx: pixel.width,
    heightPx: pixel.height,
    widthCm: (extent.cx / EMU_PER_INCH) * 2.54,
    heightCm: (extent.cy / EMU_PER_INCH) * 2.54,
  };
}
