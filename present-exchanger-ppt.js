
// Open PowerPoint
var w = new ActiveXObject("Powerpoint.Application");
w.Visible=1;

// Create new presentation file
var p = w.Presentations.Add();

// Register members.
var members = [
		"Alice",
		"Bob",
		"Carol",
		"Dan",
		"Erin"
];

// Create a copy of the array.
var copied = [];
for (var i = 0; i < members.length; i ++) {
	copied[i] = members[i];
}

// Create a shuffled list from the copied array.
// (The copied array will be empty after this operation.)
var shuffled = [];
for (var i = 0; i < members.length; i ++) {
	var rand = Math.floor(Math.random() * copied.length);
	shuffled[i] = copied.splice(rand, 1);
}

// Add a title slide.
var s = p.Slides.Add(1,1);
s.Shapes(1).TextFrame.TextRange.Text = "Christmas party!";
s.Shapes(2).TextFrame.TextRange.Text = "Igarashi Lab\n2013/12/26";

// Insert 5 blank slides.
function addBlankSlides(startPage, n) {
	var s;
	for (var page = startPage; page < startPage + n; page ++) {
		s = p.Slides.Add(page, 2);
	}
	return s;
}
var s2 = addBlankSlides(2, 5);

// Show all members.
s2.Shapes(1).TextFrame.TextRange.Text = "Members";
for (var id = 0; id < members.length; id ++) {
	s2.Shapes(2).TextFrame.TextRange.Text += members[id] +
			(id == members.length - 1 ? "" : "\n");
}

// Insert 5 blank slides.
var s3 = addBlankSlides(7, 5);

// Prepare for the results...
s3 = p.Slides.Add(12, 1);
s3.Shapes(1).TextFrame.TextRange.Text = "Results";
s3.Shapes(2).TextFrame.TextRange.Text = "will be shown in the next slides!!";

// Give instructions.
for (var i = 0; i < members.length; i ++) {
	var present = p.Slides.Add(13 + i, 2);
	present.Shapes(1).TextFrame.TextRange.Text = shuffled[i];
	present.Shapes(2).TextFrame.TextRange.Text = "Please give your present to: " + shuffled[(i + 1) % shuffled.length];
}
