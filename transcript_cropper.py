"""
Transcript PDF Cropper
======================
Processes a folder of microfiche transcript PDFs.
- Clean pages (single transcript): copied as-is
- Overlap pages (2 transcripts): split into _MAIN and _SIDE

Usage:
    TranscriptCropper.exe --input C:\\path\\to\\pdfs --output C:\\path\\to\\output
"""

import os
import sys
import argparse
import statistics
import fitz  # PyMuPDF

# ── CONFIG (tweak if crops look wrong) ──
OVERLAP_THRESHOLD = 630      # CropW above this = overlap page
CLEAN_WIDTH_FALLBACK = 570   # used if auto-detect fails


def compute_clean_width(pdf_paths):
    clean_widths = []
    for path in pdf_paths:
        try:
            doc = fitz.open(path)
            for page in doc:
                w = page.rect.width
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


def process_pdfs(input_folder, output_folder, overlap_threshold, forced_clean_width):
    os.makedirs(output_folder, exist_ok=True)
    log_path = os.path.join(output_folder, "_crop_log.txt")

    pdf_paths = sorted([
        os.path.join(input_folder, f)
        for f in os.listdir(input_folder)
        if f.lower().endswith(".pdf")
    ])

    if not pdf_paths:
        print("[ERROR] No PDF files found in input folder.")
        return

    print(f"[INFO] Found {len(pdf_paths)} PDF files.")
    print("[INFO] Scanning for clean transcript width...")

    clean_width = forced_clean_width if forced_clean_width else compute_clean_width(pdf_paths)

    log_lines = []
    log_lines.append("=== TRANSCRIPT CROP LOG ===")
    log_lines.append(f"Input folder:      {input_folder}")
    log_lines.append(f"Output folder:     {output_folder}")
    log_lines.append(f"Clean width used:  {clean_width:.1f}pt")
    log_lines.append(f"Overlap threshold: {overlap_threshold}pt")
    log_lines.append("")
    log_lines.append(f"{'File':<50} {'W':<8} {'Type':<10} Action")
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

            for page_num in range(num_pages):
                page = doc[page_num]
                rect = page.rect
                w = rect.width
                page_suffix = f"_p{page_num+1}" if num_pages > 1 else ""

                if w <= overlap_threshold:
                    # CLEAN - copy as-is
                    out_path = os.path.join(output_folder, f"{stem}{page_suffix}_CLEAN.pdf")
                    out_doc = fitz.open()
                    out_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
                    out_doc.save(out_path)
                    out_doc.close()
                    log_lines.append(f"{filename:<50} {w:<8.1f} {'CLEAN':<10} → {os.path.basename(out_path)}")
                    total_clean += 1

                else:
                    # OVERLAP - split into MAIN (left) + SIDE (right)
                    main_rect = fitz.Rect(rect.x0, rect.y0, rect.x0 + clean_width, rect.y1)
                    side_rect = fitz.Rect(rect.x0 + clean_width, rect.y0, rect.x1, rect.y1)

                    for label, crop_rect in [("MAIN", main_rect), ("SIDE", side_rect)]:
                        out_path = os.path.join(output_folder, f"{stem}{page_suffix}_{label}.pdf")
                        out_doc = fitz.open()
                        out_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
                        out_doc[0].set_cropbox(crop_rect)
                        out_doc.save(out_path)
                        out_doc.close()

                    log_lines.append(f"{filename:<50} {w:<8.1f} {'OVERLAP':<10} Split at x={clean_width:.0f}pt → MAIN + SIDE")
                    total_overlap += 1

            doc.close()

        except Exception as e:
            log_lines.append(f"{filename:<50} {'?':<8} {'ERROR':<10} {str(e)}")
            total_errors += 1

        if (idx + 1) % 100 == 0 or (idx + 1) == len(pdf_paths):
            print(f"  Progress: {idx+1}/{len(pdf_paths)} files...")

    log_lines.append("")
    log_lines.append("=== SUMMARY ===")
    log_lines.append(f"Total files:    {len(pdf_paths)}")
    log_lines.append(f"Clean pages:    {total_clean}")
    log_lines.append(f"Overlap pages:  {total_overlap}")
    log_lines.append(f"Errors:         {total_errors}")

    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))

    print("\n[DONE]")
    print(f"  Clean:    {total_clean}")
    print(f"  Overlap:  {total_overlap}")
    print(f"  Errors:   {total_errors}")
    print(f"  Log:      {log_path}")


def main():
    global OVERLAP_THRESHOLD

    parser = argparse.ArgumentParser(
        description="Transcript PDF Cropper - splits overlap microfiche pages"
    )
    parser.add_argument("--input",  "-i", required=True, help="Input folder with PDFs")
    parser.add_argument("--output", "-o", required=True, help="Output folder for cropped PDFs")
    parser.add_argument("--threshold", "-t", type=float, default=630,
                        help="Width threshold for overlap detection (default: 630)")
    parser.add_argument("--clean-width", "-c", type=float, default=None,
                        help="Force a fixed clean transcript width (skip auto-detect)")

    args = parser.parse_args()

    if not os.path.isdir(args.input):
        print(f"[ERROR] Input folder not found: {args.input}")
        sys.exit(1)

    print("=" * 50)
    print("  TRANSCRIPT PDF CROPPER")
    print("=" * 50)
    print(f"  Input:     {args.input}")
    print(f"  Output:    {args.output}")
    print(f"  Threshold: {args.threshold}pt")
    print()

    process_pdfs(args.input, args.output, args.threshold, args.clean_width)


if __name__ == "__main__":
    main()
