#!/usr/bin/env python3
"""
Find page numbers of text occurrences in a DOCX file.

This script converts the DOCX to PDF using LibreOffice, then searches
for the specified text in the PDF and returns the page numbers.
"""

import argparse
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import fitz  # PyMuPDF


def convert_docx_to_pdf(docx_path: str, output_dir: str) -> str:
    """Convert DOCX to PDF using LibreOffice headless mode."""
    docx_path = Path(docx_path).resolve()
    
    if not docx_path.exists():
        raise FileNotFoundError(f"DOCX file not found: {docx_path}")
    
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        str(docx_path),
        "--outdir", output_dir
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")
    
    pdf_name = docx_path.stem + ".pdf"
    pdf_path = Path(output_dir) / pdf_name
    
    if not pdf_path.exists():
        raise RuntimeError(f"PDF was not created: {pdf_path}")
    
    return str(pdf_path)


def find_text_in_pdf(pdf_path: str, search_text: str, case_sensitive: bool = False) -> list[dict]:
    """
    Find all occurrences of text in a PDF and return page numbers.
    
    Returns a list of dicts with:
    - page: 1-indexed page number
    - count: number of occurrences on that page
    - snippets: list of text snippets with context
    """
    doc = fitz.open(pdf_path)
    results = []
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        page_text = page.get_text()
        
        # Handle case sensitivity
        if case_sensitive:
            search_in = page_text
            search_for_text = search_text
        else:
            search_in = page_text.lower()
            search_for_text = search_text.lower()
        
        # Count occurrences and find positions
        snippets = []
        count = 0
        pos = 0
        
        while True:
            idx = search_in.find(search_for_text, pos)
            if idx == -1:
                break
            
            count += 1
            
            # Get context (50 chars before and after)
            start = max(0, idx - 50)
            end = min(len(page_text), idx + len(search_text) + 50)
            snippet = page_text[start:end].replace('\n', ' ').strip()
            snippets.append(f"...{snippet}...")
            pos = idx + 1
        
        if count > 0:
            results.append({
                "page": page_num + 1,  # 1-indexed
                "count": count,
                "snippets": snippets[:3]  # Limit to 3 snippets per page
            })
    
    doc.close()
    return results


def find_text_pages(docx_path: str, search_text: str, case_sensitive: bool = False, 
                    keep_pdf: bool = False, output_dir: str = None) -> dict:
    """
    Main function to find text occurrences in a DOCX file.
    
    Args:
        docx_path: Path to the DOCX file
        search_text: Text to search for
        case_sensitive: Whether the search should be case-sensitive
        keep_pdf: Whether to keep the generated PDF file
        output_dir: Directory for the PDF (uses temp if None)
    
    Returns:
        Dictionary with:
        - total_occurrences: Total number of occurrences
        - pages: List of page numbers where text appears
        - details: Detailed results per page
        - pdf_path: Path to PDF (if keep_pdf=True)
    """
    # Create temp directory if not specified
    if output_dir:
        temp_dir = output_dir
        cleanup = False
    else:
        temp_dir = tempfile.mkdtemp(prefix="docx_search_")
        cleanup = not keep_pdf
    
    try:
        # Convert to PDF
        pdf_path = convert_docx_to_pdf(docx_path, temp_dir)
        
        # Find text in PDF
        results = find_text_in_pdf(pdf_path, search_text, case_sensitive)
        
        # Summarize results
        total = sum(r["count"] for r in results)
        pages = [r["page"] for r in results]
        
        response = {
            "total_occurrences": total,
            "pages": pages,
            "details": results
        }
        
        if keep_pdf:
            response["pdf_path"] = pdf_path
        elif cleanup:
            os.unlink(pdf_path)
            os.rmdir(temp_dir)
        
        return response
        
    except Exception as e:
        if cleanup and os.path.exists(temp_dir):
            import shutil
            shutil.rmtree(temp_dir, ignore_errors=True)
        raise


def main():
    parser = argparse.ArgumentParser(
        description="Find page numbers of text occurrences in DOCX files"
    )
    parser.add_argument("input_path", help="Path to a DOCX file or directory containing DOCX files")
    parser.add_argument("search_text", help="Text to search for")
    parser.add_argument("-c", "--case-sensitive", action="store_true",
                        help="Make search case-sensitive")
    parser.add_argument("--keep-pdf", action="store_true",
                        help="Keep the generated PDF file")
    parser.add_argument("-o", "--output-dir", 
                        help="Directory for the PDF output")
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="Show detailed results with snippets")
    
    args = parser.parse_args()
    
    input_path = Path(args.input_path).resolve()
    
    if not input_path.exists():
        print(f"Error: Path not found: {input_path}", file=sys.stderr)
        sys.exit(1)
        
    docx_files = []
    if input_path.is_dir():
        docx_files = sorted(list(input_path.glob("*.docx")))
        if not docx_files:
            print(f"No DOCX files found in directory: {input_path}")
            return
        print(f"Processing {len(docx_files)} files in '{input_path}'...\n")
    else:
        docx_files = [input_path]
        
    total_found = 0
    files_with_matches = 0
    
    for docx_file in docx_files:
        try:
            print(f"Scanning: {docx_file.name}...", end=" ", flush=True)
            
            result = find_text_pages(
                str(docx_file),
                args.search_text,
                args.case_sensitive,
                args.keep_pdf,
                args.output_dir
            )
            
            if result["total_occurrences"] > 0:
                print(f"FOUND ({result['total_occurrences']} matches)")
                print(f"  Pages: {', '.join(map(str, result['pages']))}")
                
                if args.verbose:
                    print("  Details:")
                    for detail in result["details"]:
                        print(f"    Page {detail['page']} ({detail['count']} matches):")
                        for snippet in detail["snippets"]:
                            print(f"      {snippet}")
                print("-" * 40)
                total_found += result["total_occurrences"]
                files_with_matches += 1
            else:
                print("No matches.")
                
        except Exception as e:
            print(f"ERROR: {e}")

    print(f"\nSummary: Found {total_found} total occurrences in {files_with_matches} of {len(docx_files)} files.")

if __name__ == "__main__":
    main()
