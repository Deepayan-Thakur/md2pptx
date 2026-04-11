import os
import sys
import glob
import subprocess
import argparse

def main():
    parser = argparse.ArgumentParser(description="Run md2pptx against all test cases.")
    parser.add_argument("--template", default="template1.pptx", 
                        choices=["template.pptx", "template1.pptx", "template2.pptx"],
                        help="Template to use (default: template1.pptx)")
    args = parser.parse_args()

    test_files = glob.glob("test_cases/*.md")
    
    if not test_files:
        print("No markdown files found in test_cases/")
        sys.exit(1)
        
    print(f"Found {len(test_files)} test files. Generating presentations using '{args.template}'...")
    
    os.makedirs(os.path.join("md2pptx", "outputs"), exist_ok=True)
    
    for idx, md_file in enumerate(test_files, start=1):
        filename = os.path.basename(md_file)
        basename = os.path.splitext(filename)[0]
        output_file = os.path.join("md2pptx", "outputs", f"{basename}_{args.template.split('.')[0]}.pptx")
        
        print(f"\n--- Processing [{idx}/{len(test_files)}]: {md_file} ---")
        
        try:
            subprocess.run([
                sys.executable,
                "main.py",
                md_file,
                output_file,
                "--template",
                args.template
            ], check=True)
            print(f"-> Verification Success: Output saved to {output_file}")
        except subprocess.CalledProcessError as e:
            print(f"-> ERROR processing {md_file}: {e}")

if __name__ == "__main__":
    main()
