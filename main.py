import warnings
import os
from bs4 import BeautifulSoup
from ebooklib import epub
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import re

# Suppress warnings
warnings.filterwarnings('ignore', category=UserWarning, module='ebooklib.epub')
warnings.filterwarnings('ignore', category=FutureWarning, module='ebooklib.epub')

def epub_to_text(epub_path, output_path):
    """Convert EPUB to text file, preserving chapter structure"""
    try:
        # Read EPUB file
        book = epub.read_epub(epub_path)
        
        # Open output file
        with open(output_path, 'w', encoding='utf-8') as out_file:
            # Get all items in order
            items = list(book.get_items())
            
            # Write book title if available
            if book.get_metadata('DC', 'title'):
                title = book.get_metadata('DC', 'title')[0][0]
                out_file.write(f"# {title}\n\n")
            
            chapters_processed = 0
            processed_titles = set()  # Track processed titles to avoid duplication
            
            for item in items:
                # Process HTML content (type 9 is HTML content)
                if item.get_type() == 9:
                    print(f"Processing: {item.get_name()}")
                    # Parse HTML content
                    soup = BeautifulSoup(item.content, 'html.parser')
                    
                    # Remove script and style elements
                    for elem in soup(['script', 'style']):
                        elem.decompose()
                    
                    # Extract headings first to check for duplicates
                    headings = []
                    for elem in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6']):
                        heading_text = elem.get_text(strip=True)
                        if heading_text and heading_text not in processed_titles:
                            headings.append((elem.name, heading_text))
                            processed_titles.add(heading_text)
                    
                    # Get document text - structured approach
                    content_extracted = False
                    
                    # First write the unique headings
                    for heading_name, heading_text in headings:
                        level = int(heading_name[1])
                        out_file.write(f"\n\n{'#' * level} {heading_text}\n")
                        content_extracted = True
                    
                    # Then process other content elements
                    for elem in soup.find_all(['p', 'div', 'span', 'li', 'td', 'th', 'a', 'blockquote', 'pre', 'code']):
                        # Skip if this element is inside a heading we've already processed
                        if elem.find_parent(['h1', 'h2', 'h3', 'h4', 'h5', 'h6']):
                            continue
                            
                        text = elem.get_text(strip=True)
                        if text:  # Only process non-empty elements
                            content_extracted = True
                            if elem.name in ['li']:
                                # List items
                                out_file.write(f"\n- {text}")
                            elif elem.name in ['pre', 'code']:
                                # Preserve formatting for code blocks
                                out_file.write(f"\n```\n{elem.get_text()}\n```\n")
                            else:
                                # Regular paragraphs
                                out_file.write(f"\n{text}\n")
                    
                    # If no content was extracted with the structural approach, fallback to getting all text
                    # but exclude headings we've already processed
                    if not content_extracted:
                        # First extract all headings to exclude
                        all_headings = [h.get_text(strip=True) for h in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])]
                        
                        # Get remaining text by removing elements we've already processed
                        for h in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6']):
                            h.decompose()
                            
                        text = soup.get_text(separator='\n\n', strip=True)
                        if text:
                            out_file.write(f"\n{text}\n")
                            content_extracted = True
                    
                    if content_extracted:
                        chapters_processed += 1
            
            # Add summary at the end
            out_file.write(f"\n\n--- End of conversion ---\n")
            out_file.write(f"Processed {chapters_processed} document sections\n")
        
        print(f"Successfully converted {epub_path} to {output_path}")
        print(f"Processed {chapters_processed} document sections")
        return True
    except Exception as e:
        print(f"Error converting EPUB {os.path.basename(epub_path)}: {str(e)}")
        # Try to provide more detailed error information
        import traceback
        traceback.print_exc()
        return False

class EpubConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("EPUB to TXT Converter")
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Files list frame
        files_frame = ttk.LabelFrame(main_frame, text="EPUB Files")
        files_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Listbox for selected files
        self.files_listbox = tk.Listbox(files_frame, width=60, height=10)
        self.files_listbox.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar for listbox
        scrollbar = ttk.Scrollbar(files_frame, orient="vertical", command=self.files_listbox.yview)
        scrollbar.grid(row=0, column=2, sticky=(tk.N, tk.S))
        self.files_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Buttons for file management
        btn_frame = ttk.Frame(files_frame)
        btn_frame.grid(row=1, column=0, columnspan=3, pady=5)
        
        ttk.Button(btn_frame, text="Add Files", command=self.add_files).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="Remove Selected", command=self.remove_file).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="Clear All", command=self.clear_files).grid(row=0, column=2, padx=5)
        
        # Output directory selection
        ttk.Label(main_frame, text="Output Directory:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_dir = tk.StringVar()
        
        # Set default output path to the program's directory
        default_output_path = os.path.dirname(os.path.abspath(__file__))
        self.output_dir.set(default_output_path)
        
        ttk.Entry(main_frame, textvariable=self.output_dir, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_output_dir).grid(row=1, column=2)
        
        # Progress bar
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(progress_frame, text="Current File Progress:").grid(row=0, column=0, sticky=tk.W)
        self.progress_var = tk.IntVar()
        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=300, 
                                           mode="determinate", variable=self.progress_var)
        self.progress_bar.grid(row=0, column=1, padx=5, sticky=(tk.W, tk.E))
        
        # Convert button
        ttk.Button(main_frame, text="Convert All", command=self.convert).grid(row=3, column=1, pady=10)
        
        # Progress label
        self.status_var = tk.StringVar()
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=4, column=0, columnspan=3)
        
        # Store the list of input files
        self.input_files = []
        # Store conversion status for each file (0=pending, 1=success, -1=failed)
        self.conversion_status = {}

    def add_files(self):
        filenames = filedialog.askopenfilenames(
            title="Select EPUB files",
            filetypes=[("EPUB files", "*.epub"), ("All files", "*.*")]
        )
        if filenames:
            for file in filenames:
                if file not in self.input_files:
                    self.input_files.append(file)
                    self.files_listbox.insert(tk.END, os.path.basename(file))
                    self.conversion_status[file] = 0  # Mark as pending

    def remove_file(self):
        selected = self.files_listbox.curselection()
        if selected:
            # Remove in reverse order to avoid index shifting issues
            for index in sorted(selected, reverse=True):
                file = self.input_files[index]
                del self.input_files[index]
                self.files_listbox.delete(index)
                if file in self.conversion_status:
                    del self.conversion_status[file]

    def clear_files(self):
        self.input_files.clear()
        self.conversion_status.clear()
        self.files_listbox.delete(0, tk.END)

    def browse_output_dir(self):
        directory = filedialog.askdirectory(
            title="Select Output Directory"
        )
        if directory:
            self.output_dir.set(directory)

    def update_file_status(self, index, status):
        """Update status of file in the listbox"""
        filename = os.path.basename(self.input_files[index])
        if status == 1:  # Success
            self.files_listbox.delete(index)
            self.files_listbox.insert(index, f"✓ {filename}")
            self.files_listbox.itemconfig(index, {'fg': 'green'})
        elif status == -1:  # Failed
            self.files_listbox.delete(index)
            self.files_listbox.insert(index, f"✗ {filename}")
            self.files_listbox.itemconfig(index, {'fg': 'red'})
        self.root.update()

    def process_file_with_progress(self, input_path, output_path):
        """Process a file with comprehensive deduplication strategy"""
        try:
            # Read EPUB file
            book = epub.read_epub(input_path)
            
            # Get all items to determine total work
            items = list(book.get_items())
            
            html_items = [item for item in items if item.get_type() == 9]
            total_items = len(html_items)
            if total_items == 0:
                self.status_var.set("No HTML content found in EPUB file")
                return False
            
            # Reset progress bar
            self.progress_var.set(0)
            self.progress_bar['maximum'] = total_items
            self.root.update()
            
            # Collect all content with initial deduplication
            all_content = []
            seen_content = set()  # Track unique content during collection
            toc_sections = set()  # Special handling for TOC sections
            
            # Identify TOC sections for special handling
            for i, item in enumerate(html_items):
                self.progress_var.set(i + 1)
                self.status_var.set(f"Analyzing: {item.get_name()}")
                self.root.update()
                
                soup = BeautifulSoup(item.content, 'html.parser')
                
                # Check if this looks like a table of contents
                links = soup.find_all('a')
                if len(links) > 5:  # Arbitrary threshold for TOC detection
                    link_texts = [link.get_text().strip() for link in links]
                    if any(text.startswith('第') and ('卷' in text or '章' in text) for text in link_texts if text):
                        toc_sections.add(item.get_name())
            
            # Reset progress for second pass
            self.progress_var.set(0)
            self.root.update()
            
            # Second pass: collect all content with initial deduplication
            for i, item in enumerate(html_items):
                self.progress_var.set(i + 1)
                self.status_var.set(f"Collecting from: {item.get_name()}")
                self.root.update()
                
                # Skip TOC sections
                if item.get_name() in toc_sections:
                    continue
                    
                soup = BeautifulSoup(item.content, 'html.parser')
                
                # Remove scripts, styles, and other non-content elements
                for elem in soup(['script', 'style', 'meta', 'link', 'noscript']):
                    elem.decompose()
                
                # Extract headings
                for elem in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6']):
                    # Extract text
                    heading_text = ' '.join(elem.get_text().strip().split())
                    if not heading_text:
                        continue
                    
                    # Always include headings, even if duplicate (for structure)
                    level = int(elem.name[1])
                    all_content.append(("heading", level, heading_text))
                
                # Extract main content with inline deduplication
                for elem in soup.find_all(['p', 'div', 'span', 'li', 'td', 'th', 'a', 'blockquote', 'pre', 'code']):
                    # Skip elements inside headings
                    if elem.find_parent(['h1', 'h2', 'h3', 'h4', 'h5', 'h6']):
                        continue
                    
                    # Skip empty elements
                    text = ' '.join(elem.get_text().strip().split())
                    if not text or len(text) < 5:  # Skip very short fragments
                        continue
                    
                    # For non-headings, check if we've seen this exact content before
                    content_key = text.lower()  # Case-insensitive comparison
                    if content_key in seen_content:
                        continue  # Skip this duplicate content
                    
                    seen_content.add(content_key)
                    
                    if elem.name in ['li']:
                        all_content.append(("list", text))
                    elif elem.name in ['pre', 'code']:
                        all_content.append(("code", text))
                    else:
                        all_content.append(("paragraph", text))
            
            # Open output file and write all content
            with open(output_path, 'w', encoding='utf-8') as out_file:
                # Write book title if available
                if book.get_metadata('DC', 'title'):
                    title = book.get_metadata('DC', 'title')[0][0]
                    out_file.write(f"# {title}\n\n")
                
                # Write all content in sequence
                current_section = None
                
                for content_type, *content_data in all_content:
                    # Add section dividers between headings of level 1 or 2
                    if content_type == "heading":
                        level, text = content_data
                        if level <= 2 and current_section != text:
                            if current_section is not None:  # Not the first section
                                out_file.write("\n\n" + "-" * 40 + "\n\n")
                            current_section = text
                        
                        out_file.write(f"\n\n{'#' * level} {text}\n")
                    elif content_type == "list":
                        out_file.write(f"\n- {content_data[0]}")
                    elif content_type == "code":
                        out_file.write(f"\n```\n{content_data[0]}\n```\n")
                    elif content_type == "paragraph":
                        out_file.write(f"\n{content_data[0]}\n")
            
            # After writing the output file, perform advanced post-processing
            self.status_var.set("Post-processing: advanced deduplication...")
            self.root.update()
            
            # Read the file we just wrote
            with open(output_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Split the content into paragraphs
            paragraphs = re.split(r'\n\n+', content)
            
            # STAGE 1: Remove exact duplicates (case insensitive)
            unique_paragraphs = []
            seen_paragraphs = set()
            
            for para in paragraphs:
                para_stripped = para.strip()
                # Always keep headings and section dividers
                if para_stripped.startswith('#') or para_stripped.startswith('-' * 10):
                    unique_paragraphs.append(para)
                    continue
                    
                # Skip empty paragraphs
                if not para_stripped or len(para_stripped) < 10:
                    continue
                    
                # Check for exact duplicates
                para_key = para_stripped.lower()
                if para_key not in seen_paragraphs:
                    seen_paragraphs.add(para_key)
                    unique_paragraphs.append(para)
            
            # STAGE 2: Check for contained paragraphs
            paragraphs_to_keep = [True] * len(unique_paragraphs)
            
            for i in range(len(unique_paragraphs)):
                para_i = unique_paragraphs[i].strip()
                # Skip special paragraphs
                if para_i.startswith('#') or para_i.startswith('-' * 10) or len(para_i) < 10:
                    continue
                    
                # Skip if already marked for removal
                if not paragraphs_to_keep[i]:
                    continue
                    
                for j in range(len(unique_paragraphs)):
                    if i == j or not paragraphs_to_keep[i]:
                        continue
                        
                    para_j = unique_paragraphs[j].strip()
                    # Skip special paragraphs for comparison
                    if para_j.startswith('#') or para_j.startswith('-' * 10):
                        continue
                        
                    # Check containment
                    if para_i in para_j and para_i != para_j:
                        paragraphs_to_keep[i] = False
                        break
            
            # Create final filtered paragraphs
            final_paragraphs = [p for i, p in enumerate(unique_paragraphs) if paragraphs_to_keep[i]]
            
            # Write the filtered content back to the file
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write('\n\n'.join(final_paragraphs))
            
            total_removed = len(paragraphs) - len(final_paragraphs)
            self.status_var.set(f"Completed with deduplication. Removed {total_removed} duplicate paragraphs.")
            return True
        except Exception as e:
            error_msg = f"Error converting EPUB {os.path.basename(input_path)}: {str(e)}"
            print(error_msg)
            
            # Print to console
            import traceback
            traceback.print_exc()
            return False

    def convert(self):
        if not self.input_files:
            messagebox.showerror("Error", "Please add at least one EPUB file!")
            return
        
        output_dir = self.output_dir.get()
        if not output_dir:
            messagebox.showerror("Error", "Please select an output directory!")
            return
            
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create output directory: {str(e)}")
                return
        
        # Initialize progress tracking
        total_files = len(self.input_files)
        successful = 0
        failed = 0
        
        for i, input_path in enumerate(self.input_files):
            # Update status
            self.status_var.set(f"Converting file {i+1} of {total_files}: {os.path.basename(input_path)}")
            self.root.update()
            
            # Generate output path
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(output_dir, f"{base_name}.txt")
            
            # Convert the file with progress indication
            if self.process_file_with_progress(input_path, output_path):
                successful += 1
                self.conversion_status[input_path] = 1
                self.update_file_status(i, 1)
            else:
                failed += 1
                self.conversion_status[input_path] = -1
                self.update_file_status(i, -1)
        
        # Reset progress bar
        self.progress_var.set(0)
        
        # Show completion message
        if failed == 0:
            self.status_var.set(f"All {successful} files converted successfully!")
            messagebox.showinfo("Success", f"All {successful} files have been converted to {output_dir}")
        else:
            self.status_var.set(f"Completed with {successful} successes and {failed} failures")
            messagebox.showwarning("Partial Success", 
                                  f"Converted {successful} files successfully.\n{failed} files failed to convert.")

def main():
    root = tk.Tk()
    app = EpubConverterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()