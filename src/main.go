package main

import (
	"bytes"
	"context"
	"fmt"
	"image"
	"image/jpeg"
	_ "image/png"
	"io"
	"log"
	"net"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"time"

	"baliance.com/gooxml/common"
	"baliance.com/gooxml/measurement"
	"baliance.com/gooxml/presentation"
	"github.com/chromedp/chromedp"
	"github.com/hooklift/iso9660"
)

const (
	contentDir = "../content"
	tempDir    = "temp_extracted"
	outputPPTX = "../output.pptx"
)

func main() {
	log.Println("Starting Storyline to PPTX converter...")

	// 1. Find and Extract ISO
	isoPath, err := findISO(contentDir)
	if err != nil {
		log.Fatalf("Error finding ISO: %v", err)
	}
	log.Printf("Found ISO: %s", isoPath)

	if err := extractISO(isoPath, tempDir); err != nil {
		log.Fatalf("Error extracting ISO: %v", err)
	}
	defer os.RemoveAll(tempDir) // Cleanup

	// 2. Start Local Server
	port, err := getFreePort()
	if err != nil {
		log.Fatalf("Error getting free port: %v", err)
	}
	server := &http.Server{Addr: fmt.Sprintf(":%d", port), Handler: http.FileServer(http.Dir(tempDir))}
	go func() {
		if err := server.ListenAndServe(); err != nil && err != http.ErrServerClosed {
			log.Fatalf("HTTP server error: %v", err)
		}
	}()
	defer server.Close()
	baseURL := fmt.Sprintf("http://localhost:%d/story.html", port)
	log.Printf("Serving content at %s", baseURL)

	// 3. Setup Chromedp
	opts := append(chromedp.DefaultExecAllocatorOptions[:],
		chromedp.WindowSize(1280, 720),
	)
	allocCtx, cancel := chromedp.NewExecAllocator(context.Background(), opts...)
	defer cancel()
	ctx, cancel := chromedp.NewContext(allocCtx)
	defer cancel()

	// 4. Initialize PPTX
	ppt := presentation.New()
	defer func() {
		f, err := os.Create(outputPPTX)
		if err != nil {
			log.Printf("Error creating output file: %v", err)
			return
		}
		defer f.Close()
		if err := ppt.Save(f); err != nil {
			log.Printf("Error saving PPT: %v", err)
		}
	}()

	// 5. Scrape Loop
	log.Println("Navigating to story...")
	if err := chromedp.Run(ctx, chromedp.Navigate(baseURL)); err != nil {
		log.Fatalf("Error navigating: %v", err)
	}

	// Wait for initial load
	time.Sleep(5 * time.Second)

	// Inject CSS to hide controls (slider, play buttons)
	hideControlsJS := `
		const style = document.createElement('style');
		style.innerHTML = '.controls-group, .cs-controls, .area-primary { display: none !important; }';
		document.head.appendChild(style);
	`
	if err := chromedp.Run(ctx, chromedp.Evaluate(hideControlsJS, nil)); err != nil {
		log.Printf("Warning: Could not hide controls: %v", err)
	}

	slideIndex := 1
	for {
		log.Printf("Processing Slide %d...", slideIndex)

		// Wait for slide content to settle
		time.Sleep(2 * time.Second)

		// Capture Screenshot
		var buf []byte
		if err := chromedp.Run(ctx, chromedp.CaptureScreenshot(&buf)); err != nil {
			log.Printf("Error capturing screenshot: %v", err)
			break
		}

		// Extract Text (for editable notes)
		var slideText string
		extractTextJS := `document.body.innerText`
		if err := chromedp.Run(ctx, chromedp.Evaluate(extractTextJS, &slideText)); err != nil {
			log.Printf("Warning: Could not extract text: %v", err)
		}

		// Add to PPTX
		if err := addSlideToPPT(ppt, buf, slideText); err != nil {
			log.Printf("Error adding slide to PPT: %v", err)
		}

		// Check for "Next" button and click
		var nextDisabled bool
		// Common Storyline Next Button Selectors
		checkNextJS := `
			(function() {
				const btn = document.querySelector('#next') || 
							document.querySelector('.next-button') || 
							document.querySelector('div[data-model-id="5hW..."]'); // Generic fallback
				if (!btn) return true; // No button found, maybe end
				if (btn.classList.contains('disabled') || btn.getAttribute('aria-disabled') === 'true') return true;
				btn.click();
				return false;
			})()
		`
		if err := chromedp.Run(ctx, chromedp.Evaluate(checkNextJS, &nextDisabled)); err != nil {
			log.Printf("Error checking next button: %v", err)
			break
		}

		if nextDisabled {
			log.Println("End of presentation reached.")
			break
		}

		slideIndex++
		if slideIndex > 100 { // Safety break
			log.Println("Max slides reached, stopping.")
			break
		}
	}

	log.Printf("Saved PowerPoint to %s", outputPPTX)
}

// addSlideToPPT adds a screenshot and notes to a new slide
func addSlideToPPT(ppt *presentation.Presentation, imgBytes []byte, notes string) error {
	slide := ppt.AddSlide()

	// Add Image
	// Decode PNG from memory to ensure validity and convert to JPEG
	// Converting to JPEG avoids potential PNG decoding issues in gooxml v1.0.1
	srcImg, _, err := image.Decode(bytes.NewReader(imgBytes))
	if err != nil {
		return fmt.Errorf("failed to decode screenshot: %w", err)
	}

	// Write image to temp file
	tmpFile, err := os.CreateTemp(tempDir, "slide-*.jpg")
	if err != nil {
		return err
	}

	if err := jpeg.Encode(tmpFile, srcImg, &jpeg.Options{Quality: 90}); err != nil {
		tmpFile.Close()
		return err
	}
	tmpFile.Sync() // Ensure data is flushed to disk
	if err := tmpFile.Close(); err != nil {
		return err
	}

	absPath, err := filepath.Abs(tmpFile.Name())
	if err != nil {
		return err
	}

	// Verify image is valid before passing to gooxml
	// This helps diagnose "image must have a valid size" errors
	f, err := os.Open(absPath)
	if err != nil {
		return fmt.Errorf("unable to open image for verification: %w", err)
	}
	fi, err := f.Stat()
	if err != nil {
		f.Close()
		return fmt.Errorf("unable to stat image: %w", err)
	}
	var cfg image.Config
	cfg, _, err = image.DecodeConfig(f)
	f.Close()
	if err != nil {
		return fmt.Errorf("invalid image data: %w", err)
	}
	log.Printf("Debug: Image verified at %s (Size: %d bytes, Dim: %dx%d)", absPath, fi.Size(), cfg.Width, cfg.Height)

	img := common.Image{
		Path:   absPath,
		Format: "jpeg",
	}
	imgRef, err := ppt.AddImage(img)
	if err != nil {
		return err
	}

	// Create an image box filling the slide (assuming 16:9 aspect ratio roughly)
	imgBox := slide.AddImage(imgRef)
	imgBox.Properties().SetPosition(0, 0)
	// Standard PPT size is often 16x9 inches or similar, gooxml defaults might vary.
	// We set it to cover a standard wide slide.
	imgBox.Properties().SetSize(measurement.Distance(13.33*measurement.Inch), measurement.Distance(7.5*measurement.Inch))

	// Add Notes (Editable Text)
	// Note: gooxml support for notes is limited in older versions,
	// so we might just print it to console or try to add a hidden text box if notes fail.
	// For this example, we will add a text box at the bottom (off-screen or visible) containing the text.

	// Adding a text box with the extracted text for maintainability
	tb := slide.AddTextBox()
	tb.Properties().SetPosition(measurement.Distance(0.5*measurement.Inch), measurement.Distance(7.6*measurement.Inch))
	tb.Properties().SetSize(measurement.Distance(12*measurement.Inch), measurement.Distance(2*measurement.Inch))
	p := tb.AddParagraph()
	run := p.AddRun()
	run.SetText("Extracted Text: " + strings.ReplaceAll(notes, "\n", " "))
	run.Properties().SetSize(10 * measurement.Point)

	return nil
}

// findISO looks for the first .iso file in the directory
func findISO(dir string) (string, error) {
	var isoPath string
	err := filepath.WalkDir(dir, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if !d.IsDir() && strings.HasSuffix(strings.ToLower(d.Name()), ".iso") {
			isoPath = path
			return io.EOF // Stop search
		}
		return nil
	})
	if err == io.EOF {
		return isoPath, nil
	}
	if isoPath == "" {
		return "", fmt.Errorf("no ISO file found in %s", dir)
	}
	return isoPath, err
}

// extractISO extracts the ISO content to dest
func extractISO(isoPath, dest string) error {
	f, err := os.Open(isoPath)
	if err != nil {
		return err
	}
	defer f.Close()

	r, err := iso9660.NewReader(f)
	if err != nil {
		return fmt.Errorf("failed to open ISO reader: %w", err)
	}

	for {
		f, err := r.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		// Construct target path
		// Note: iso9660 paths are usually / separated and uppercase
		relPath := strings.TrimLeft(f.Name(), "/")
		targetPath := filepath.Join(dest, relPath)

		if f.IsDir() {
			if err := os.MkdirAll(targetPath, 0755); err != nil {
				return err
			}
			continue
		}

		// Ensure parent dir exists
		if err := os.MkdirAll(filepath.Dir(targetPath), 0755); err != nil {
			return err
		}

		// Write file
		outFile, err := os.Create(targetPath)
		if err != nil {
			return err
		}
		if _, err := io.Copy(outFile, f.Sys().(io.Reader)); err != nil {
			outFile.Close()
			return err
		}
		outFile.Close()
	}
	return nil
}

// getFreePort asks the kernel for a free open port
func getFreePort() (int, error) {
	addr, err := net.ResolveTCPAddr("tcp", "localhost:0")
	if err != nil {
		return 0, err
	}
	l, err := net.ListenTCP("tcp", addr)
	if err != nil {
		return 0, err
	}
	defer l.Close()
	return l.Addr().(*net.TCPAddr).Port, nil
}

// Helper to execute JS and ignore errors if needed
func eval(ctx context.Context, js string) {
	var res interface{}
	_ = chromedp.Run(ctx, chromedp.Evaluate(js, &res))
}
