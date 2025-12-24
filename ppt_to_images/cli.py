"""
Command-line interface for ppt-to-images.
"""

import argparse
import json
import sys
import logging
from pathlib import Path

from .converter import PPTConverter, ConversionError


def setup_logging(verbose: bool = False):
    """Setup logging configuration."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )


def main():
    """Main entry point for CLI."""
    parser = argparse.ArgumentParser(
        prog="ppt-to-images",
        description="Convert PPT/PPTX/PDF files to image sequences"
    )
    
    # Required arguments
    parser.add_argument(
        "input",
        help="Input file path (PPT/PPTX/PDF)"
    )
    
    # Output options
    parser.add_argument(
        "-o", "--output-dir",
        help="Output directory for image files"
    )
    
    parser.add_argument(
        "-f", "--format",
        choices=["file", "base64", "both"],
        default="file",
        help="Output format (default: file)"
    )
    
    # Conversion options
    parser.add_argument(
        "--dpi",
        type=int,
        default=200,
        help="Image DPI/resolution (default: 200)"
    )
    
    parser.add_argument(
        "--extract-text",
        action="store_true",
        help="Extract text content from slides"
    )
    
    # Output formatting
    parser.add_argument(
        "--output-json",
        action="store_true",
        help="Output result as JSON"
    )
    
    # Advanced options
    parser.add_argument(
        "--temp-dir",
        help="Custom temporary directory"
    )
    
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="Keep temporary files (for debugging)"
    )
    
    # Logging
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose output"
    )
    
    parser.add_argument(
        "--version",
        action="version",
        version="%(prog)s 0.1.0"
    )
    
    args = parser.parse_args()
    
    # Setup logging
    setup_logging(args.verbose)
    logger = logging.getLogger(__name__)
    
    # Validate arguments
    if args.format in ("file", "both") and not args.output_dir:
        parser.error(f"--output-dir is required for format '{args.format}'")
    
    # Check input file exists
    input_path = Path(args.input)
    if not input_path.exists():
        logger.error(f"Input file not found: {input_path}")
        sys.exit(1)
    
    # Check supported file type
    ext = input_path.suffix.lower()
    if ext not in (".ppt", ".pptx", ".pdf"):
        logger.error(f"Unsupported file type: {ext}")
        logger.error("Supported types: .ppt, .pptx, .pdf")
        sys.exit(1)
    
    try:
        # Create converter
        converter = PPTConverter(
            dpi=args.dpi,
            cleanup_temp=not args.keep_temp,
            temp_dir=args.temp_dir
        )
        
        # Run conversion
        logger.info(f"Converting: {input_path}")
        result = converter.convert(
            input_path=input_path,
            output_dir=args.output_dir,
            output_format=args.format,
            extract_text=args.extract_text
        )
        
        # Output result
        if args.output_json:
            print(json.dumps(result, indent=2, ensure_ascii=False))
        else:
            print(f"Converted {result['count']} pages")
            if args.format in ("file", "both"):
                print(f"Output directory: {args.output_dir}")
                for path in result["images"]:
                    print(f"  - {path}")
            if args.format == "base64":
                print("Base64 output:")
                for i, b64 in enumerate(result["images"], 1):
                    print(f"  Page {i}: {b64[:50]}..." if len(b64) > 50 else f"  Page {i}: {b64}")
            if args.extract_text and result.get("texts"):
                print("\nExtracted text:")
                for i, text in enumerate(result["texts"], 1):
                    preview = text[:100] + "..." if len(text) > 100 else text
                    print(f"  Page {i}: {preview}")
        
        logger.info("Conversion completed successfully")
        
    except ConversionError as e:
        logger.error(f"Conversion failed: {e}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()

