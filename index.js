const officegen = require('officegen')
const fs = require('fs')

// Create an empty PowerPoint object:
let pptx = officegen('pptx')

// Officegen calling this function after finishing to generate the pptx document:
pptx.on('finalize', function(written) {
  console.log(
    'Finish to create a Microsoft PowerPoint document.'
  )
})

// Officegen calling this function to report errors:
pptx.on('error', function(err) {
  console.log(err)
})

// Let's add a title slide:

let slide = pptx.makeTitleSlide('Officegen', 'Example to a PowerPoint document')

// Pie chart slide example:

slide = pptx.makeNewSlide()
slide.name = 'Pie Chart slide'
slide.back = 'ffff00'
slide.addChart(
  {
    title: 'My production',
    renderType: 'pie',
    data:
	[
      {
        name: 'Oil',
        labels: ['Czech Republic', 'Ireland', 'Germany', 'Australia', 'Austria', 'UK', 'Belgium'],
        values: [301, 201, 165, 139, 128,  99, 60],
        colors: ['ff0000', '00ff00', '0000ff', 'ffff00', 'ff00ff', '00ffff', '000000']
      }
    ]
  }
)

// Let's generate the PowerPoint document into a file:

let out = fs.createWriteStream('example.pptx')

out.on('error', function(err) {
  console.log(err)
})

// Async call to generate the output file:
pptx.generate(out)