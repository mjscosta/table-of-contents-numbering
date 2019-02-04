
/**
 * Function that returns a node as a dictionary, given a set of parameters.
 * The dictionary has the form {h:h, ch:ch, l:l}.
 * @param h - heading.
 * @param ch - children.
 * @param l - level.
 */
function Node(h, ch, l) {
  return {
    h: h,
    ch: ch,
    l: l
  };
}

function Heading(heading, level) {
  return {h: heading, l:level};
}

/**
 * Heading Tree class.
 */
var HeadingTree = function(headingList) {
  this.root = {h:null, ch:[], l:0};
  
  /**
   * Add an heading to the tree at the specified level.
   *
   * @param root - tree root
   */
  this.addHeading_ = function(node, heading, level) {
    if (node.l + 1 == level) {
      if (node.ch.length == 0) {
        node.ch.push(null);
      }
      node.ch.push(Node(heading, [], level));
    } else if (node.l + 1 < level) {
      if (node.ch.length == 0) {
        node.ch.push(this.addHeading_(Node(null, [], node.l + 1), heading, level));
      } else {
        this.addHeading_(node.ch[node.ch.length - 1], heading, level);
      }
    }
    return node;
  }
  
  this.addHeading = function(heading, level) {
    this.addHeading_(this.root, heading, level);
  }
  
  
  /**
   * Walk the tree from the given, and apply the processor to all nodes
   * providing the path to the processor.
   * @param node - the current node of the tree, to walk.
   * @param path - the list of indexes of nodes in order, from the root till the current node.
   * @param context - processor context, to allow for configuration parameters.
   * @param processor - a function of the form processor(path, node, result, context), that performs an 
   * action, and returns the result, in the result parameter, that can be a list.
   * @param result - result parameter as an accumulator, parameter, e.g. a list
   * or a reference object.
   */
  this.walkTree_ = function(node, path, processor, result, context) {
      if (node != null) {
        if (node.h != null) {
          processor(path, node, result, context);
        }
        
        for (var i = 0; i < node.ch.length; i++) {
          var new_path = path.slice();
          new_path.push(i);
          this.walkTree_(node.ch[i], new_path, processor, result, context);
        }
      }  
  }
  
  /**
   * Walk the full tree and apply the processor to all nodes
   * providing the path to the processor, returning the result in 
   * the result parameter as an accumulator, parameter, e.g. a list
   * or a reference object.
   *
   * @param processor - a function of the form processor(path, node, result), that performs an 
   * action, and returns the result, in the result parameter, that can be a list.
   * @param result - result parameter as an accumulator, parameter, e.g. a list
   * or a reference object.
   * @param context - processor context, to allow for configuration parameters.
   */
  this.walkTree = function(processor, result, context) {
    this.walkTree_(this.root, [], processor, result, context);
  }
  
  if(headingList) {
    for(var i in headingList) {
      this.addHeading(headingList[i].h, headingList[i].l);
    }
  }

}



/**
 * Returns a list of paragraph elements with their respective level, with heading in {1 .. maxHeading}, respecting the order they ocurr
 * in the document. Example: [{h:Paragraph, l:1} ...].
 *
 * @param maxHeadingLevel - the maximum heading level that will be returned.
 */
function findParagraphsWithHeading(maxHeadingLevel) {
  var paragraphElements = [];

  // Get the body section of the active document.
  var body = DocumentApp.getActiveDocument().getBody();
  
  // Define the search parameters.
  var searchType = DocumentApp.ElementType.PARAGRAPH;
  var searchResult = null;
  
  // Get the heading values from the enum, in order to compare them
  // enum ParagraphHeading = [Normal, Heading 1, Heading 2, Heading 3, Heading 4, Heading 5, Heading 6, Title, Subtitle]
  var headings = DocumentApp.ParagraphHeading.values() 
  // Search until the paragraph is found.
  while (searchResult = body.findElement(searchType, searchResult)) {
    var paragraph = searchResult.getElement().asParagraph();
    for (var headingLevel = 1; headingLevel <= maxHeadingLevel; headingLevel++) {
      if (paragraph.getHeading() == headings[headingLevel]) {
        paragraphElements.push(Heading(paragraph,headingLevel));
        break;
      }
    }
  }
  
  return paragraphElements;
}

/**
 * Cleans the heading numbering of the given paragraph.
 *
 * @param paragraph - the paragraph, to clean the numbering.
 */
function resetHeading(paragraph) {
  paragraph.setText(paragraph.getText().replace(/^\d(\.?\d)*(\.? |\.)/gm, ""));
}


/**
 * Sets the paragraph numbering, according to the path, and format.
 *
 * @param path - list of indexes of the children nodes till the current 
 * "node" or heading.
 * @param paragraph - the document paragraph element, to be modified.
 * @param format - an indication of how to format the numbering of the headings.
 * At the moment, it only supports one numbering scheme.
 */
function setParagraphNumbering(path, paragraph, format) {
  const numberingRegEx = /^\d(\.?\d)*(\.? |\.)/gm;
  var numbering = "";
  
  for(var i = 0; i < path.length -1; i++) {
    numbering += path[i] + ".";
  }

  numbering += path[path.length -1] + " ";
  
  
  if(paragraph.getText().match(numberingRegEx)) {
    paragraph.setText(paragraph.getText().replace(numberingRegEx, numbering));
  } else {
    paragraph.setText(numbering + paragraph.getText())
  }
}

/**
 * Processor function that clears or sets the numbering format of the headings, 
 * depending on the format value.
 *
 * @see HeadingTree.walkTree(...)
 * param path - the list of indexes of the children till the current node.
 * @param node - the current node being processed.
 * @param result - not used.
 * @param format - the format to apply for the number formating, one of {'none', 'format_1'}.
 */
function updateHeadingProcessor(path, node, result, format) {
  if (node != null) {
    if (node.h != null) {
      if(format == 'none') {
        resetHeading(node.h);
      } else {
        setParagraphNumbering(path, node.h, format);
      }
    }
  }
}

/**
 * Update all headings according to the maxLevel and format.
 */
function updateHeadings(maxHeadingLevel, format) {
  // retrieve list of paragraphs. 
  var headingParagraphs = findParagraphsWithHeading(maxHeadingLevel);
  var headingTree = new HeadingTree(headingParagraphs);
  
  headingTree.walkTree(updateHeadingProcessor, null, format);

}

function updateHeadings1(maxHeadingLevel, format) {
  updateHeadings(6,'format_1');
}

/* ==================================== Test Section ============================================ */

/**
 * Creates a string from a tree.
 */
function nodeToStringProcessor(path, node, result) {
  var nodeResult = "";
  
  if (node != null) {
    if (node.h != null) {
      for(var i = 0; i < path.length -1; i++) {
        nodeResult += path[i] + ".";
      }
      
      nodeResult += path[path.length -1];
      nodeResult += " " + node.h;
      
      result.push(nodeResult);
    }
  }
}

/**
 * Given a list of headings in the format {h:heading, l:level}, creates a tree of headings, and walks the 
 * tree, processing each heading. The result of that processing is a list with a string per heading, 
 * containing, the heading numbering (tree path), and the heading test. 
 *
 * Example:
 * 
 * testTree([Heading("H1",1), [Heading("H2",2)], ["1 H1", "1.1 H2"]), should succeed.
 *
 * If the expected result and output do not match, then a dialog with an error is displayed.
 */
function testTree(headingList, expectedOutput, testId, outputHeadings) {
  var headingTree = new HeadingTree(headingList);
  
  var headingsWithNumbering = [];
  headingTree.walkTree(nodeToStringProcessor, headingsWithNumbering);
  
  var ui = DocumentApp.getUi();
  
  if(headingsWithNumbering.length != expectedOutput.length) {
    ui.alert("Test Error: " + testId, "Actual Length: " + headingsWithNumbering.length + " Expected Length: " + expectedOutput.length, ui.ButtonSet.OK);
  } else {
    for(var i in headingsWithNumbering) {
      if(headingsWithNumbering[i] != expectedOutput[i]) {
        ui.alert("Test Error: " + testId, "Actual: " + headingsWithNumbering[i] + " Expected: " + expectedOutput[i], ui.ButtonSet.OK);
      }
    }
  }
  
  if(outputHeadings) {
    DocumentApp.getActiveDocument().getBody().appendParagraph("Start Test: " + testId + " ------------");
    for(var i in headingsWithNumbering) {
      DocumentApp.getActiveDocument().getBody().appendParagraph(headingsWithNumbering[i])
    }
    DocumentApp.getActiveDocument().getBody().appendParagraph("End Test: " + testId + " ------------");
  }
}

function test1() {
  var printOutput = true;
  
  var headingList1 = [Heading("H1",1), Heading("H1",1), Heading("H1",1)];
  var expected1 = ["1 H1", "2 H1", "3 H1"];
  testTree(headingList1, expected1, "test1", printOutput);
  
  var headingList2 = [Heading("H2",2), Heading("H1",1), Heading("H1",1)];
  var expected2 = ["0.1 H2", "1 H1", "2 H1"];
  testTree(headingList2, expected2, "test2", printOutput);
  
  var headingList3 = [Heading("H1",1), Heading("H2",2), Heading("H3",3)];
  var expected3 = ["1 H1", "1.1 H2", "1.1.1 H3"];
  testTree(headingList3, expected3, "test3", printOutput);
  
  var headingList4 = [Heading("H1",1), Heading("H3",3)];
  var expected4 = ["1 H1", "1.0.1 H3"];
  testTree(headingList4, expected4, "test4", printOutput);
}

/* ================================== End Test Section ========================================== */


/** Application Settings and Interaction **/
function onOpen() {
  DocumentApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .createMenu('Table Of Contents Numbering')
    .addItem('Show Settings', 'showSettingsSideBar')
    .addToUi();
}

function showSettingsSideBar() {
  var html = HtmlService.createHtmlOutputFromFile('Settings')
    .setTitle('Settings')
    .setWidth(300);
  DocumentApp.getUi()
    .showSidebar(html);
}
