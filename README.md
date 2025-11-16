#!/usr/bin/env python3
"""
Standalone Word Document Endnote to Incipit Converter
Converts traditional endnotes to incipit format with dynamic PAGEREF fields
Author: Eric Caplan
Date: November 2024
"""

import xml.dom.minidom as minidom
import zipfile
import re
import os
import sys
import shutil
from pathlib import Path

def unpack_docx(docx_path, extract_dir):
    """Extract a .docx file to a directory"""
    print(f"Unpacking {docx_path}...")
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

def pack_docx(source_dir, output_path):
    """Pack a directory back into a .docx file"""
    print(f"Creating {output_path}...")
    
    # Create a new zip file
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Walk through all files in the source directory
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                file_path = os.path.join(root, file)
                # Calculate the archive name (relative path from source_dir)
                arcname = os.path.relpath(file_path, source_dir)
                zipf.write(file_path, arcname)

def extract_endnotes(endnotes_path):
    """Extract endnote content from endnotes.xml"""
    print("Extracting endnotes...")
    with open(endnotes_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    dom = minidom.parseString(content)
    endnotes = {}
    
    for endnote in dom.getElementsByTagName('w:endnote'):
        endnote_id = endnote.getAttribute('w:id')
        if endnote_id and endnote_id not in ['-1', '0']:  # Skip separators
            # Extract text content from the endnote
            text_elements = []
            for t_elem in endnote.getElementsByTagName('w:t'):
                text_elements.append(t_elem.firstChild.nodeValue if t_elem.firstChild else '')
            
            # Join text and clean up
            full_text = ''.join(text_elements)
            # Remove the endnote number at the start
            full_text = re.sub(r'^\s*\d*\s*', '', full_text).strip()
            
            endnotes[endnote_id] = full_text
    
    return endnotes

def process_endnote_references(doc_path, output_path):
    """Replace endnote references with bookmarks and extract context"""
    print("Processing endnote references and creating bookmarks...")
    with open(doc_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    dom = minidom.parseString(content)
    references = {}
    bookmark_id = 1000  # Start bookmark IDs at 1000 to avoid conflicts
    
    # Process all paragraphs to find endnote references
    for para in dom.getElementsByTagName('w:p'):
        # Extract paragraph text for context
        para_text = []
        endnote_runs = []
        
        # Look for text and endnote references
        for run in para.getElementsByTagName('w:r'):
            # Check for endnote reference
            endnote_refs = run.getElementsByTagName('w:endnoteReference')
            if endnote_refs:
                endnote_id = endnote_refs[0].getAttribute('w:id')
                endnote_runs.append((run, endnote_id))
                # Mark the position
                para_text.append('[ENDNOTE]')
            else:
                # Get text
                for t_elem in run.getElementsByTagName('w:t'):
                    if t_elem.firstChild:
                        para_text.append(t_elem.firstChild.nodeValue)
        
        # Join paragraph text
        full_para = ''.join(para_text)
        
        # Process endnote references
        for run, endnote_id in endnote_runs:
            # Find text before the endnote marker
            parts = full_para.split('[ENDNOTE]')
            if parts[0]:
                # Get the last few words before the endnote
                words_before = parts[0].strip().split()
                if len(words_before) >= 3:
                    first_three = ' '.join(words_before[-3:])
                else:
                    first_three = ' '.join(words_before) if words_before else "Beginning of section"
                
                # Clean up the first three words
                first_three = re.sub(r'[,;.!?]+$', '', first_three)
                
                # Create bookmark name (must be valid XML name)
                bookmark_name = f"endnote_{endnote_id}"
                
                references[endnote_id] = {
                    'bookmark': bookmark_name,
                    'bookmark_id': str(bookmark_id),
                    'first_three': first_three
                }
                
                # Replace the endnote reference run with a bookmark
                parent = run.parentNode
                
                # Create bookmark start element
                bookmark_start = dom.createElement('w:bookmarkStart')
                bookmark_start.setAttribute('w:id', str(bookmark_id))
                bookmark_start.setAttribute('w:name', bookmark_name)
                
                # Create bookmark end element
                bookmark_end = dom.createElement('w:bookmarkEnd')
                bookmark_end.setAttribute('w:id', str(bookmark_id))
                
                # Insert bookmark elements
                parent.insertBefore(bookmark_start, run)
                parent.insertBefore(bookmark_end, run)
                
                # Remove the endnote reference run
                parent.removeChild(run)
                
                bookmark_id += 1
    
    # Save the modified document
    with open(output_path, 'w', encoding='utf-8') as f:
        # Pretty print for easier debugging if needed
        dom.writexml(f, encoding='utf-8')
    
    return references

def create_notes_section_xml(endnotes, references):
    """Create the XML for the Notes section with PAGEREF fields"""
    notes_xml = []
    
    # Add a page break
    notes_xml.append('''  <w:p>
    <w:pPr>
      <w:pageBreakBefore/>
    </w:pPr>
  </w:p>''')
    
    # Add Notes heading
    notes_xml.append('''  <w:p>
    <w:pPr>
      <w:pStyle w:val="Heading1"/>
    </w:pPr>
    <w:r>
      <w:t>Notes</w:t>
    </w:r>
  </w:p>''')
    
    # Add each incipit note with PAGEREF field
    for note_id in sorted(endnotes.keys(), key=lambda x: int(x)):
        citation = endnotes[note_id]
        
        # Escape special XML characters
        citation = citation.replace('&', '&amp;')
        citation = citation.replace('<', '&lt;')
        citation = citation.replace('>', '&gt;')
        citation = citation.replace('"', '&quot;')
        
        if note_id in references:
            ref = references[note_id]
            bookmark_name = ref['bookmark']
            first_three = ref['first_three']
            
            # Escape first_three too
            first_three = first_three.replace('&', '&amp;')
            first_three = first_three.replace('<', '&lt;')
            first_three = first_three.replace('>', '&gt;')
            first_three = first_three.replace('"', '&quot;')
            
            # Create note with PAGEREF field for page number
            note_xml = f'''  <w:p>
    <w:pPr>
      <w:spacing w:after="120"/>
    </w:pPr>
    <w:r>
      <w:fldSimple w:instr=" PAGEREF {bookmark_name} \\h ">
        <w:r>
          <w:t>[Page]</w:t>
        </w:r>
      </w:fldSimple>
    </w:r>
    <w:r>
      <w:t xml:space="preserve"> </w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:i/>
        <w:iCs/>
      </w:rPr>
      <w:t>{first_three}:</w:t>
    </w:r>
    <w:r>
      <w:t xml:space="preserve"> {citation}</w:t>
    </w:r>
  </w:p>'''
        else:
            # No reference found - use placeholder
            citation_escaped = citation.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            note_xml = f'''  <w:p>
    <w:pPr>
      <w:spacing w:after="120"/>
    </w:pPr>
    <w:r>
      <w:t>[Missing reference] {citation_escaped}</w:t>
    </w:r>
  </w:p>'''
        
        notes_xml.append(note_xml)
    
    return '\n'.join(notes_xml)

def add_notes_to_document(doc_path, notes_xml, output_path):
    """Add the notes section to the document"""
    with open(doc_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Find the closing body tag
    body_close_pos = content.rfind('</w:body>')
    
    if body_close_pos == -1:
        print("Warning: Could not find closing body tag")
        return False
    
    # Find the last sectPr (section properties) element
    sect_pr_pos = content.rfind('<w:sectPr', 0, body_close_pos)
    
    if sect_pr_pos != -1:
        # Insert notes before the section properties
        insert_pos = sect_pr_pos
    else:
        # Insert before the closing body tag
        insert_pos = body_close_pos
    
    # Insert the notes
    new_content = content[:insert_pos] + notes_xml + '\n' + content[insert_pos:]
    
    # Save the modified document
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    return True

def convert_document(input_docx, output_docx=None):
    """Main conversion function"""
    # Set up paths
    input_path = Path(input_docx)
    if not input_path.exists():
        print(f"Error: Input file '{input_docx}' not found")
        return False
    
    if output_docx is None:
        output_docx = input_path.stem + '_incipit.docx'
    
    output_path = Path(output_docx)
    
    # Create temporary directory for extraction
    temp_dir = Path('temp_docx_extract')
    if temp_dir.exists():
        shutil.rmtree(temp_dir)
    temp_dir.mkdir()
    
    try:
        print("\n" + "="*60)
        print("ENDNOTE TO INCIPIT CONVERTER")
        print("="*60)
        
        # Step 1: Unpack the document
        unpack_docx(input_path, temp_dir)
        
        # Step 2: Check if document has endnotes
        endnotes_file = temp_dir / 'word' / 'endnotes.xml'
        if not endnotes_file.exists():
            print("No endnotes found in this document.")
            return False
        
        # Step 3: Extract endnotes
        endnotes = extract_endnotes(endnotes_file)
        print(f"Found {len(endnotes)} endnotes")
        
        # Step 4: Process document - replace references with bookmarks
        doc_file = temp_dir / 'word' / 'document.xml'
        doc_temp = temp_dir / 'word' / 'document_temp.xml'
        references = process_endnote_references(doc_file, doc_temp)
        print(f"Created {len(references)} bookmarks")
        
        # Step 5: Create notes section
        print("Creating Notes section with PAGEREF fields...")
        notes_xml = create_notes_section_xml(endnotes, references)
        
        # Step 6: Add notes to document
        success = add_notes_to_document(doc_temp, notes_xml, doc_file)
        if not success:
            print("Error: Failed to add notes section")
            return False
        
        # Remove temporary file
        doc_temp.unlink()
        
        # Step 7: Pack the document back
        pack_docx(temp_dir, output_path)
        
        print("\n" + "="*60)
        print("✓ CONVERSION COMPLETE!")
        print("="*60)
        print(f"\nOutput file: {output_path}")
        print("\n📝 IMPORTANT INSTRUCTIONS:")
        print("After opening the document in Microsoft Word:")
        print("  1. Press Command+A (Select All)")
        print("  2. Press F9 (Update Fields)")
        print("  3. Page numbers will update from '[Page]' to actual numbers")
        print("\n💡 TIP: Repeat Command+A then F9 whenever you change formatting")
        print("="*60 + "\n")
        
        return True
        
    except Exception as e:
        print(f"Error during conversion: {e}")
        import traceback
        traceback.print_exc()
        return False
        
    finally:
        # Clean up temporary directory
        if temp_dir.exists():
            shutil.rmtree(temp_dir)

def main():
    """Main entry point for command-line usage"""
    print("\nWORD DOCUMENT ENDNOTE TO INCIPIT CONVERTER")
    print("Converts traditional superscript endnotes to incipit format")
    print("with dynamic page references (PAGEREF fields)\n")
    
    if len(sys.argv) < 2:
        print("Usage:")
        print(f"  python {sys.argv[0]} input.docx [output.docx]")
        print("\nExample:")
        print(f"  python {sys.argv[0]} manuscript.docx")
        print(f"  python {sys.argv[0]} manuscript.docx manuscript_incipit.docx")
        print("\nIf no output file is specified, '_incipit' will be added to the input filename")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    success = convert_document(input_file, output_file)
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()
