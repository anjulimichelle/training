// Adds a "Trainees" menu to the Classroom Training Assessment doc:
//   Add trainee                      - copies Trainee 01's section (with its Team and
//                                      rating dropdowns) to the end of the doc, renumbers
//                                      it, and resets the other fields to placeholders.
//   Copy Trainee 01 dropdowns to all - puts Trainee 01's Team and Performance Rating
//                                      lines into every section that still shows the
//                                      "[Team name]" or "[Select rating]" placeholder.

var NAVY = '#263F70';
var MUTED = '#6B7280';
var PLACEHOLDER = '#9CA3AF';

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Trainees')
    .addItem('Add trainee', 'addTrainee')
    .addItem('Copy Trainee 01 dropdowns to all', 'copyDropdownsToAll')
    .addToUi();
}

function isTraineeHeading_(el) {
  return el.getType() === DocumentApp.ElementType.PARAGRAPH &&
    el.asParagraph().getHeading() === DocumentApp.ParagraphHeading.HEADING3 &&
    /^Trainee \d+/.test(el.asParagraph().getText());
}

function isBlankParagraph_(el) {
  return el.getType() === DocumentApp.ElementType.PARAGRAPH && el.asParagraph().getText().trim() === '';
}

// Each section runs from its heading up to the divider above the next heading (or the doc end).
function getSections_(body) {
  var heads = [];
  for (var i = 0; i < body.getNumChildren(); i++) {
    if (isTraineeHeading_(body.getChild(i))) heads.push(i);
  }
  return heads.map(function (h, n) {
    var end = body.getNumChildren();
    if (n + 1 < heads.length) {
      end = heads[n + 1];
      while (end - 1 > h && isBlankParagraph_(body.getChild(end - 1))) end--;
    }
    return { start: h, end: end };
  });
}

function findLine_(body, section, prefix) {
  for (var i = section.start; i < section.end; i++) {
    var el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH && el.asParagraph().getText().indexOf(prefix) === 0) return i;
  }
  return -1;
}

function style_(p, color, size, bold) {
  var t = p.editAsText();
  if (t.getText().length === 0) return;
  t.setForegroundColor(color).setFontSize(size).setBold(bold).setFontFamily('Arial');
}

function copyDropdownsToAll() {
  var body = DocumentApp.getActiveDocument().getBody();
  var sections = getSections_(body);
  if (sections.length < 2) {
    DocumentApp.getUi().alert('Only one trainee section found.');
    return;
  }
  var lines = [['Team:', '[Team name]'], ['Performance Rating:', '[Select rating]']];
  var templates = lines.map(function (l) {
    var i = findLine_(body, sections[0], l[0]);
    return i < 0 ? null : body.getChild(i).copy();
  });
  var replaced = 0;
  // Work from the bottom up so earlier indexes stay valid.
  for (var s = sections.length - 1; s >= 1; s--) {
    lines.forEach(function (l, k) {
      if (!templates[k]) return;
      var i = findLine_(body, sections[s], l[0]);
      if (i < 0 || body.getChild(i).asParagraph().getText().indexOf(l[1]) < 0) return;
      body.insertParagraph(i, templates[k].copy());
      body.getChild(i + 1).removeFromParent();
      replaced++;
    });
  }
  DocumentApp.getUi().alert('Updated ' + replaced + ' line(s). Pick each trainee\'s team and rating in their dropdowns.');
}

function addTrainee() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var sections = getSections_(body);
  if (!sections.length) {
    DocumentApp.getUi().alert('No "Trainee NN" heading found to copy.');
    return;
  }

  var copies = [];
  for (var j = sections[0].start; j < sections[0].end; j++) copies.push(body.getChild(j).copy());

  body.appendParagraph('').appendHorizontalRule();

  var added = [];
  copies.forEach(function (c) {
    var type = c.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      // Appending a table also adds an empty paragraph after it; drop that one.
      var prev = body.getChild(body.getNumChildren() - 1);
      var p = body.appendParagraph(c);
      var beforePrev = body.getChildIndex(prev) > 0 ? body.getChild(body.getChildIndex(prev) - 1) : null;
      if (isBlankParagraph_(prev) && beforePrev && beforePrev.getType() === DocumentApp.ElementType.TABLE) {
        prev.removeFromParent();
      }
      added.push(p);
    } else if (type === DocumentApp.ElementType.TABLE) {
      added.push(body.appendTable(c));
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      added.push(body.appendListItem(c));
    }
  });

  var num = sections.length + 1;
  var label = (num < 10 ? '0' : '') + num;
  var mode = null;
  var firstInBlock = true;
  var heading = null;
  var extras = [];

  added.forEach(function (el) {
    if (el.getType() === DocumentApp.ElementType.TABLE) {
      resetStats_(el.asTable());
      return;
    }
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) {
      if (mode) extras.push(el);
      return;
    }
    var p = el.asParagraph();
    var text = p.getText();

    if (isTraineeHeading_(p)) {
      p.setText('Trainee ' + label + ' — [Trainee Name]');
      style_(p, NAVY, 11.5, true);
      heading = p;
    } else if (text === "Trainer's Justification") {
      mode = 'justification';
      firstInBlock = true;
    } else if (text === 'Areas for Improvement') {
      mode = 'areas';
      firstInBlock = true;
    } else if (mode) {
      if (firstInBlock) {
        p.setText(mode === 'justification'
          ? '[Enter justification]'
          : '[Enter coaching opportunities and recommended improvements]');
        style_(p, PLACEHOLDER, 10, false);
        firstInBlock = false;
      } else {
        extras.push(p);
      }
    }
    // The Team and Performance Rating lines keep Trainee 01's dropdowns; pick the new values there.
  });

  extras.forEach(function (el) {
    try {
      el.removeFromParent();
    } catch (e) {
      // The doc's last paragraph can't be removed, so empty it instead.
      if (el.getType() === DocumentApp.ElementType.PARAGRAPH) el.asParagraph().setText('');
    }
  });

  if (heading) doc.setCursor(doc.newPosition(heading, 0));
}

function resetStats_(table) {
  for (var r = 0; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    for (var c = 0; c < row.getNumCells(); c++) {
      var cell = row.getCell(c);
      for (var k = cell.getNumChildren() - 1; k >= 1; k--) {
        var child = cell.getChild(k);
        if (k === 1 && child.getType() === DocumentApp.ElementType.PARAGRAPH) {
          child.asParagraph().setText('—');
          style_(child.asParagraph(), PLACEHOLDER, 10, false);
        } else {
          child.removeFromParent();
        }
      }
    }
  }
}
