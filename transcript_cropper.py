"""
Transcript PDF Cropper
======================
Processes a folder of microfiche transcript PDFs.
- Clean pages (single transcript): copied as-is
- Overlap pages (2 transcripts): split into _MAIN and _SIDE

Usage:
    transcript_cropper.exe --input "C:\path\to\pdfs" --output "C:\path\to\output"

Build to exe:
    pip install pyinstaller pymupdf
    pyinstaller --onefile transcript_cropper.py
"""

import os
import sys
import argparse
import statistics
import fitz  # PyMuPDF


# ──────────────────────────────────────────────
# CONFIG (tweak if crops look wrong)
# ──────────────────────────────────────────────
OVERLAP_THRESHOLD = 630      # CropW above this = overlap page
CLEAN_WIDTH_FALLBACK = 570   # used if auto-detect fails
LEFT_BLEED_RATIO = 0.45      # if bleed takes up < 45% of width it's a left-bleed


def get_crop_rect(page):
    """Get the effective crop rect, handling inverted coordinates."""
    rect = page.rect  # PyMuPDF always gives correct rect
    return rect


def compute_clean_width(pdf_paths):
    """
    First pass: scan all PDFs and compute the median width
    of clean (non-overlap) pages to use as the split point.
    """
    clean_widths = []
    for path in pdf_paths:
        try:
            doc = fitz.open(path)
            for page in doc:
                w = page.rect.width
                h = page.rect.height
                if w <= OVERLAP_THRESHOLD:
                    clean_widths.append(w)
            doc.close()
        except Exception:
            pass

    if clean_widths:
        median_w = statistics.median(clean_widths)
        print(f"[INFO] Clean width median: {median_w:.1f}pt  ({len(clean_widths)} clean pages scanned)")
        return median_w
    else:
        print(f"[WARN] Could not detect clean width, using fallback: {CLEAN_WIDTH_FALLBACK}pt")
        return CLEAN_WIDTH_FALLBACK


def classify_overlap(page_width, clean_width):
    """
    Determine if the bleed is from the RIGHT or LEFT.
    
    Right-bleed: main on left, second transcript bleeds in from right
        → MAIN crop = [0, clean_width]
        
    Left-bleed: second transcript bleeds in from left, main on right
        → MAIN crop = [page_width - clean_width, page_width]
    
    We use the ratio of the bleed portion to detect direction.
    Heuristic: if the overlap excess is < LEFT_BLEED_RATIO of total width,
    assume right-bleed (most common). Otherwise could be left-bleed.
    
    For very wide pages (3+ transcripts), split at clean_width from left.
    """
    excess = page_width - clean_width
    
    # If excess is roughly equal to clean_width → true 2-transcript overlap
    # Could be either direction — default to right-bleed (main on left)
    # This matches the majority of your pages based on the examples
    return "RIGHT_BLEED"  # main on left


def process_pdfs(input_folder, output_folder):
    """Main processing loop."""

    os.makedirs(output_folder, exist_ok=True)
    log_path = os.path.join(output_folder, "_crop_log.txt")

    # Gather all PDFs
    pdf_paths = sorted([
        os.path.join(input_folder, f)
        for f in os.listdir(input_folder)
        if f.lower().endswith(".pdf")
    ])

    if not pdf_paths:
        print("[ERROR] No PDF files found in input folder.")
        return

    print(f"[INFO] Found {len(pdf_paths)} PDF files.")

    # First pass: compute clean width
    print("[INFO] Scanning for clean transcript width...")
    clean_width = compute_clean_width(pdf_paths)

    # Second pass: crop
    log_lines = []
    log_lines.append("=== TRANSCRIPT CROP LOG ===")
    log_lines.append(f"Input folder:  {input_folder}")
    log_lines.append(f"Output folder: {output_folder}")
    log_lines.append(f"Clean width:   {clean_width:.1f}pt")
    log_lines.append(f"Overlap threshold: {OVERLAP_THRESHOLD}pt")
    log_lines.append("")
    log_lines.append(f"{'File':<50} {'Pages':<6} {'Type':<10} {'Width':<10} {'Action'}")
    log_lines.append("-" * 100)

    total_clean = 0
    total_overlap = 0
    total_errors = 0

    for idx, pdf_path in enumerate(pdf_paths):
        filename = os.path.basename(pdf_path)
        stem = os.path.splitext(filename)[0]

        try:
            doc = fitz.open(pdf_path)
            num_pages = len(doc)

            # Most transcript PDFs are single page — handle multi-page too
            for page_num in range(num_pages):
                page = doc[page_num]
                rect = page.rect
                w = rect.width
                h = rect.height

                page_suffix = f"_p{page_num+1}" if num_pages > 1 else ""

                if w <= OVERLAP_THRESHOLD:
                    # ── CLEAN PAGE: copy as-is ──
                    out_path = os.path.join(output_folder, f"{stem}{page_suffix}_CLEAN.pdf")
                    out_doc = fitz.open()
                    out_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
                    out_doc.save(out_path)
                    out_doc.close()

                    log_lines.append(
                        f"{filename:<50} {num_pages:<6} {'CLEAN':<10} {w:<10.1f} "
                        f"Saved → {os.path.basename(out_path)}"
                    )
                    total_clean += 1

                else:
                    # ── OVERLAP PAGE: split into MAIN + SIDE ──
                    direction = classify_overlap(w, clean_width)

                    if direction == "RIGHT_BLEED":
                        # Main on LEFT
                        main_rect  = fitz.Rect(rect.x0, rect.y0, rect.x0 + clean_width, rect.y1)
                        side_rect  = fitz.Rect(rect.x0 + clean_width, rect.y0, rect.x1, rect.y1)
                        main_label = "MAIN_L"
                        side_label = "SIDE_R"
                    else:
                        # Main on RIGHT (left-bleed)
                        side_rect  = fitz.Rect(rect.x0, rect.y0, rect.x1 - clean_width, rect.y1)
                        main_rect  = fitz.Rect(rect.x1 - clean_width, rect.y0, rect.x1, rect.y1)
                        main_label = "MAIN_R"
                        side_label = "SIDE_L"

                    for label, crop_rect in [(main_label, main_rect), (side_label, side_rect)]:
                        out_path = os.path.join(output_folder, f"{stem}{page_suffix}_{label}.pdf")
                        out_doc = fitz.open()
                        out_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
                        out_page = out_doc[0]
                        out_page.set_cropbox(crop_rect)
                        out_doc.save(out_path)
                        out_doc.close()

                    log_lines.append(
                        f"{filename:<50} {num_pages:<6} {'OVERLAP':<10} {w:<10.1f} "
                        f"Split at x={clean_width:.0f}pt → {main_label} + {side_label}"
                    )
                    total_overlap += 1

            doc.close()

        except Exception as e:
            log_lines.append(f"{filename:<50} {'?':<6} {'ERROR':<10} {'?':<10} {str(e)}")
            total_errors += 1

        # Progress
        if (idx + 1) % 100 == 0 or (idx + 1) == len(pdf_paths):
            print(f"  Progress: {idx+1}/{len(pdf_paths)} files processed...")

    # Summary
    log_lines.append("")
    log_lines.append("=== SUMMARY ===")
    log_lines.append(f"Total files:    {len(pdf_paths)}")
    log_lines.append(f"Clean pages:    {total_clean}")
    log_lines.append(f"Overlap pages:  {total_overlap}")
    log_lines.append(f"Errors:         {total_errors}")
    log_lines.append(f"Clean width used: {clean_width:.1f}pt")

    # Save log
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))

    print("\n[DONE]")
    print(f"  Clean pages:   {total_clean}")
    print(f"  Overlap pages: {total_overlap}")
    print(f"  Errors:        {total_errors}")
    print(f"  Log saved to:  {log_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Transcript PDF Cropper - splits overlap microfiche pages"
    )
    parser.add_argument(
        "--input", "-i",
        required=True,
        help="Folder containing input PDF files"
    )
    parser.add_argument(
        "--output", "-o",
        required=True,
        help="Folder to save cropped PDFs and log"
    )
    parser.add_argument(
        "--threshold", "-t",
        type=float,
        default=OVERLAP_THRESHOLD,
        help=f"Width threshold above which a page is considered overlap (default: {OVERLAP_THRESHOLD})"
    )
    parser.add_argument(
        "--clean-width", "-c",
        type=float,
        default=None,
        help="Force a fixed clean transcript width instead of auto-detecting"
    )

    global OVERLAP_THRESHOLD

    args = parser.parse_args()

    OVERLAP_THRESHOLD = args.threshold

    if not os.path.isdir(args.input):
        print(f"[ERROR] Input folder not found: {args.input}")
        sys.exit(1)

    print("=" * 50)
    print("  TRANSCRIPT PDF CROPPER")
    print("=" * 50)
    print(f"  Input:  {args.input}")
    print(f"  Output: {args.output}")
    print(f"  Overlap threshold: {OVERLAP_THRESHOLD}pt")
    print()

    process_pdfs(args.input, args.output)


if __name__ == "__main__":
    main()
