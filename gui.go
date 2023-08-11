package main

import (
	"fmt"
	"strconv"
	"strings"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/widget"
	"gitee.com/moipa/pptx"
)

func slideString(n int) string {
	return fmt.Sprintf("ppt/slides/slide%d.xml", n)
}

func DeleteFromMeta(meta string, texted int) string {
	pre, rawPost, _ := strings.Cut(meta, "<p:sldIdLst><")
	list, post, _ := strings.Cut(rawPost, "></p:sldIdLst>")
	slides := strings.Split(list, "/><")
	cutSlides := append(slides[:texted+1], slides[len(slides)-1])
	return pre + "<p:sldIdLst><" + strings.Join(cutSlides, "/><") + "></p:sldIdLst>" + post
}

func genPptx(text string) string {
	pres, err := pptx.ReadPowerPoint("./pres.pptx")
	if err != nil {
		panic(err)
	}
	texts := strings.Split(text, "\n\n")
	for i, el := range texts {
		pres.ReplaceSlideContent("slide"+strconv.Itoa(i+2), el, i+1)
	}
	pres.Presentation = DeleteFromMeta(pres.Presentation, len(texts))
	firstRow := strings.Split(texts[0], "\n")[0]
	name := strings.Join(strings.Split(firstRow, " "), "_") + ".pptx"
	pres.WriteToFile("./" + name)
	return name
}

func main() {
	myApp := app.New()
	myWindow := myApp.NewWindow("PPTXGEN")

	input := widget.NewEntry()
	input.MultiLine = true
	input.SetPlaceHolder("Enter text...")

	label := widget.NewLabel("Output file")
	button := widget.NewButton("Generate", func() {
		label.SetText(genPptx(input.Text))
	})
	bottom := container.NewVBox(button, label)

	content := container.NewBorder(nil, bottom, nil, nil, input)
	myWindow.Resize(fyne.NewSize(640, 480))
	myWindow.SetContent(content)
	myWindow.ShowAndRun()
}
