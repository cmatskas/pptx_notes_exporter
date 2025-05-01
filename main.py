from pptx import Presentation
import os

def is_slide_hidden(slide):
    """Returns True if slide is marked as hidden in PowerPoint."""
    return slide._element.get('show', '1') == '0'

def export_powerpoint_notes(ppt_path, output_path=None):
    """
    Exports notes from each PowerPoint slide with <SLIDE> tags.
    
    Args:
        ppt_path (str): Path to the PowerPoint file (e.g., 'presentation.pptx').
        output_path (str, optional): Custom output file path. Defaults to Desktop.
    """
    # Load the presentation
    prs = Presentation(ppt_path)
    notes_content = []
    
    # Extract notes from each slide
    for slide in enumerate(prs.slides):
        if is_slide_hidden(slide):
            continue  # Skip hidden slides
        
        notes_text = "(No notes)"
        
        # Try to get notes (works if notes exist)
        if slide.has_notes_slide:
            notes_text = slide.notes_slide.notes_text_frame.text.strip()
        
        # Format with <SLIDE> tag
        notes_content.append(f"<SLIDE>\n{notes_text}\n")
    
    # Default output to Desktop if no path provided
    if not output_path:
        desktop = os.path.expanduser("~/Desktop")
        output_path = os.path.join(desktop, "Slide_Notes.txt")
    
    # Write to file
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(notes_content))
    
    print(f"Notes exported to: {output_path}")

# Example usage
if __name__ == "__main__":
    ppt_file = input("Enter PowerPoint file path (e.g., 'MyPresentation.pptx'): ").strip('"')
    export_powerpoint_notes(ppt_file)