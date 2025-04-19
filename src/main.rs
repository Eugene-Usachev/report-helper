use std::fs;
use std::ops::Add;
use std::path::Path;

use docx_rs::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let src_dir = Path::new("../src");

    fn visit_dirs(dir: &Path, mut doc: Docx) -> std::io::Result<Docx> {
        for entry in fs::read_dir(dir)? {
            let entry = entry?;
            let path = entry.path();

            if path.is_dir() {
                doc = visit_dirs(&path, doc)?;
            } else if path.is_file() {
                let path_run = Run::new()
                    .add_break(BreakType::TextWrapping)
                    .add_text(path.display().to_string().add("\n"))
                    .fonts(RunFonts::new().ascii("Times New Roman"))
                    .size(28);
                doc = doc.add_paragraph(Paragraph::new().add_run(path_run));

                let content = fs::read_to_string(&path)?;
                let mut paragraph = Paragraph::new();
                for line in content.lines() {
                    let mut run = Run::new()
                        .fonts(RunFonts::new().ascii("Courier New"))
                        .size(20);
                    for part in line.split('\t') {
                        run = run.add_text(part);
                        run = run.add_tab();
                    }
                    paragraph = paragraph
                        .add_run(run)
                        .add_run(Run::new().add_break(BreakType::TextWrapping));
                }
                doc = doc.add_paragraph(paragraph);
            }
        }

        Ok(doc)
    }

    let doc = visit_dirs(src_dir, Docx::new())?;

    let mut file = fs::File::create("output.docx")?;

    doc.build().pack(&mut file)?;

    println!("Written to output.docx");

    Ok(())
}
