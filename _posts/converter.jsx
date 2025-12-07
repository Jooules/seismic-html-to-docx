import React, { useState, useCallback, useMemo } from 'react';
import { FileText, Download, Eye, Trash2, AlertCircle, CheckCircle, ClipboardPaste, CheckSquare, Square, List } from 'lucide-react';

// Parse Seismic HTML into structured content
const parseSeismicHtml = (html) => {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');
  const content = [];
  
  // Find all widget containers
  const widgets = doc.querySelectorAll('[data-testid]');
  
  // First pass: identify accordions and their children
  const accordionMap = new Map(); // accordion element -> { title, level, children: [] }
  const widgetToAccordion = new Map(); // child widget -> parent accordion element
  
  widgets.forEach((widget) => {
    const testId = widget.getAttribute('data-testid');
    
    if (testId === 'page.accordion') {
      // Try multiple selectors for the title - some accordions have nested <p> tags
      let title = 'Untitled Section';
      
      // First try the virtual text span (cleanest source)
      const virtualTextEl = widget.querySelector('.seismic-page-divider-view-text-virtual');
      if (virtualTextEl && virtualTextEl.textContent.trim()) {
        title = virtualTextEl.textContent.trim();
      } else {
        // Fallback to the divider view text
        const titleEl = widget.querySelector('.seismic-page-divider-view-text');
        if (titleEl && titleEl.textContent.trim()) {
          title = titleEl.textContent.trim();
        }
      }
      
      const headerEl = widget.querySelector('.seismic-page-divider-view');
      let level = 2;
      if (headerEl) {
        if (headerEl.classList.contains('__heading1')) level = 1;
        else if (headerEl.classList.contains('__heading2')) level = 2;
        else if (headerEl.classList.contains('__heading3')) level = 3;
      }
      
      accordionMap.set(widget, { title, level, children: [] });
    } else {
      // Check if this widget is inside an accordion
      const parentAccordion = widget.closest('[data-testid="page.accordion"]');
      if (parentAccordion && accordionMap.has(parentAccordion)) {
        widgetToAccordion.set(widget, parentAccordion);
      }
    }
  });
  
  // Second pass: process all widgets
  widgets.forEach((widget, idx) => {
    const testId = widget.getAttribute('data-testid');
    
    // Skip non-widget testids
    if (testId === 'seismic-page-richContent') return;
    
    const parentAccordion = widgetToAccordion.get(widget);
    
    if (testId === 'page.paragraph') {
      const richContent = widget.querySelector('.seismic-page-RichTextView-content');
      if (richContent) {
        const parsed = parseParagraphContent(richContent);
        
        if (parentAccordion) {
          // Add to accordion children
          accordionMap.get(parentAccordion).children.push({ type: 'paragraph', id: `${idx}`, ...parsed });
        } else {
          // Add to top-level content
          content.push({ type: 'paragraph', id: idx, ...parsed });
        }
      }
    } 
    else if (testId === 'page.table') {
      const table = widget.querySelector('table');
      if (table) {
        const tableData = parseTable(table);
        
        if (parentAccordion) {
          accordionMap.get(parentAccordion).children.push({ type: 'table', id: `${idx}`, table: tableData });
        } else {
          content.push({ type: 'table', id: idx, table: tableData });
        }
      }
    }
    else if (testId === 'page.accordion') {
      const accordionData = accordionMap.get(widget);
      console.log('Adding accordion:', accordionData.title, 'with', accordionData.children.length, 'children');
      content.push({ type: 'accordion', id: idx, title: accordionData.title, level: accordionData.level, children: accordionData.children });
    }
    else if (testId === 'page.divider') {
      const titleEl = widget.querySelector('.seismic-page-divider-view-text');
      const title = titleEl ? titleEl.textContent.trim() : '';
      
      const headerEl = widget.querySelector('.seismic-page-divider-view');
      let level = 2; // Default to H2
      if (headerEl) {
        if (headerEl.classList.contains('__heading1')) level = 1;
        else if (headerEl.classList.contains('__heading3')) level = 3;
      }
      
      content.push({ type: 'divider', id: idx, title, level });
    }
  });
  
  return content;
};

// Parse paragraph content, detecting headings by font size and capturing all text
const parseParagraphContent = (element) => {
  const results = [];
  const tempDiv = document.createElement('div');
  tempDiv.innerHTML = element.innerHTML;
  
  // Track which text nodes we've already captured (to avoid duplicates)
  const capturedTexts = new Set();
  
  console.log('=== parseParagraphContent START ===');
  
  // Step 1a: Find native heading elements (h1, h2, h3)
  const headingData = [];
  tempDiv.querySelectorAll('h1, h2, h3').forEach(heading => {
    const tagName = heading.tagName.toLowerCase();
    const text = heading.textContent.trim();
    
    if (!text) return;
    
    console.log('Found native heading:', tagName, text.substring(0, 40));
    headingData.push({ type: tagName, content: text, element: heading });
    capturedTexts.add(text);
  });
  
  // Step 1b: Find heading spans (by font-size) and extract them with their position
  tempDiv.querySelectorAll('span[style*="font-size"]').forEach(span => {
    const fontSize = span.style.fontSize;
    const size = parseInt(fontSize);
    const text = span.textContent.trim();
    
    if (!text) return;
    
    // Skip if already captured as a native heading
    if (capturedTexts.has(text)) return;
    
    if (size >= 28) {
      headingData.push({ type: 'h1', content: text, element: span });
      capturedTexts.add(text);
    } else if (size >= 20) {
      headingData.push({ type: 'h2', content: text, element: span });
      capturedTexts.add(text);
    } else if (size >= 16) {
      const isBold = span.querySelector('b, strong') || span.closest('b, strong') || span.innerHTML.includes('<b>');
      if (isBold) {
        headingData.push({ type: 'h3', content: text, element: span });
        capturedTexts.add(text);
      }
    }
  });
  
  // Step 2: Find all list items with nesting level
  const listItems = [];
  tempDiv.querySelectorAll('li').forEach(li => {
    // Calculate nesting level by counting parent ul/ol elements
    let level = 0;
    let parent = li.parentElement;
    while (parent) {
      if (parent.tagName === 'UL' || parent.tagName === 'OL') {
        level++;
      }
      parent = parent.parentElement;
    }
    
    // Get text content, excluding nested list content
    const liClone = li.cloneNode(true);
    liClone.querySelectorAll('ul, ol').forEach(el => el.remove());
    const text = liClone.textContent.trim();
    
    if (text && !capturedTexts.has(text)) {
      // Get HTML content preserving inline formatting
      let htmlContent = liClone.innerHTML;
      htmlContent = htmlContent
        .replace(/<span[^>]*>\s*<\/span>/gi, '')
        .replace(/<span[^>]*>(.*?)<\/span>/gi, '$1')
        .replace(/<br\s*\/?>/gi, '')
        .replace(/&nbsp;/gi, ' ')
        .replace(/\s+/g, ' ')
        .trim();
      
      // Remove link tags but keep text
      htmlContent = htmlContent.replace(/<a[^>]*>(.*?)<\/a>/gi, '$1');
      // Remove underline tags
      htmlContent = htmlContent.replace(/<\/?u>/gi, '');
      
      const hasInlineFormatting = /<(b|strong|i|em)[^>]*>/i.test(htmlContent);
      
      listItems.push({ 
        type: 'listitem', 
        content: text, 
        htmlContent: hasInlineFormatting ? htmlContent : null,
        level: level,
        element: li 
      });
      capturedTexts.add(text);
    }
  });
  
  // Step 3: Find all divs and extract text content (excluding heading content)
  const textBlocks = [];
  const allDivs = tempDiv.querySelectorAll('div');
  
  console.log('Step 3: Found', allDivs.length, 'divs to process');
  
  allDivs.forEach((div, idx) => {
    // Skip if inside a list
    if (div.closest('ul, ol')) {
      console.log('Div', idx, ': SKIP - in list');
      return;
    }
    
    // Skip if inside a heading tag or anchor container (button areas)
    if (div.closest('h1, h2, h3, h4, h5, h6, .seismic-page-anchor-container')) {
      console.log('Div', idx, ': SKIP - inside heading/anchor');
      return;
    }
    
    // Check if this div has direct child divs (not counting those inside headings)
    // Only look at immediate children, not all descendants
    const directChildDivs = Array.from(div.children).filter(child => {
      if (child.tagName !== 'DIV') return false;
      // Ignore divs that are inside headings
      if (child.closest('h1, h2, h3, h4, h5, h6, .seismic-page-anchor-container')) return false;
      return true;
    });
    if (directChildDivs.length > 0) {
      console.log('Div', idx, ': SKIP - has', directChildDivs.length, 'direct child divs');
      return;
    }
    
    // Get text content, but exclude text from nested headings, lists, and child divs
    // Clone the div and remove these elements before getting text
    const clone = div.cloneNode(true);
    clone.querySelectorAll('h1, h2, h3, h4, h5, h6, .seismic-page-anchor-container, ul, ol, div').forEach(el => el.remove());
    let text = clone.textContent;
    
    // Check for nbsp-only content (used as paragraph separator)
    // \u00A0 is the nbsp character, \s matches whitespace
    const isNbspOnly = /^[\s\u00A0]*$/.test(text);
    text = text.trim();
    
    // Check if this is a break div (no meaningful text content, or nbsp-only)
    if (!text || text.length === 0 || isNbspOnly) {
      // Add break if div contains a br element OR is nbsp-only (paragraph separator)
      const hasBr = clone.querySelector('br');
      if (hasBr || isNbspOnly) {
        console.log('Div', idx, ': ADD as break', isNbspOnly ? '(nbsp)' : '(br)');
        textBlocks.push({ type: 'break', content: '', element: div });
      } else {
        console.log('Div', idx, ': SKIP - empty, no br/nbsp');
      }
      return;
    }
    
    // Skip already captured
    if (capturedTexts.has(text)) {
      console.log('Div', idx, ': SKIP - exact match already captured');
      return;
    }
    
    // Skip if this text is a substring of an already captured heading (catches heading duplicates)
    // But DON'T skip if the text is longer (body text that mentions a heading)
    let isSubstring = false;
    for (const captured of capturedTexts) {
      // Only skip if: the captured heading fully contains this text (this text is part of heading)
      // NOT if: this text contains the heading (body text mentioning heading)
      if (captured.includes(text) && text.length <= captured.length) {
        isSubstring = true;
        console.log('Div', idx, ': SKIP - substring of captured heading');
        break;
      }
    }
    if (isSubstring) return;
    
    // Get the HTML content, preserving inline formatting (b, i, strong, em)
    // Clone and clean up the div to get just the formatted text
    const htmlClone = div.cloneNode(true);
    htmlClone.querySelectorAll('h1, h2, h3, h4, h5, h6, .seismic-page-anchor-container, ul, ol, div').forEach(el => el.remove());
    
    // Get innerHTML and clean it up
    let htmlContent = htmlClone.innerHTML;
    
    // Remove empty spans and unnecessary tags, but keep b, i, strong, em
    htmlContent = htmlContent
      .replace(/<span[^>]*>\s*<\/span>/gi, '') // Remove empty spans
      .replace(/<span[^>]*>(.*?)<\/span>/gi, '$1') // Remove span wrappers but keep content
      .replace(/<br\s*\/?>/gi, '') // Remove br tags (handled separately)
      .replace(/&nbsp;/gi, ' ') // Convert nbsp to space
      .replace(/\s+/g, ' ') // Collapse whitespace
      .trim();
    
    // Check if content has any actual formatting tags
    const hasInlineFormatting = /<(b|strong|i|em)[^>]*>/i.test(htmlContent);
    
    console.log('Div', idx, ': ADD as text (hasFormatting:', hasInlineFormatting, ') -', text.substring(0, 50));
    
    // Get font size if present
    let fontSize = null;
    const fontSpan = div.querySelector('span[style*="font-size"]');
    if (fontSpan) {
      const match = fontSpan.style.fontSize.match(/(\d+)/);
      if (match) fontSize = parseInt(match[1]);
    }
    
    // Store both plain text (for deduplication) and HTML content (for rendering)
    textBlocks.push({ 
      type: 'text', 
      content: text, 
      htmlContent: hasInlineFormatting ? htmlContent : null,
      element: div, 
      fontSize 
    });
    capturedTexts.add(text);
  });
  
  console.log('Step 3 complete: Found', textBlocks.length, 'text blocks');
  
  // Step 3b: Capture text from top-level inline elements (i, b, span, etc.) not inside divs
  // These are direct children of the content container
  const inlineElements = [];
  Array.from(tempDiv.childNodes).forEach((node, idx) => {
    // Skip div elements (already processed)
    if (node.nodeType === Node.ELEMENT_NODE && node.tagName === 'DIV') return;
    
    // Skip heading elements (already processed)
    if (node.nodeType === Node.ELEMENT_NODE && /^H[1-6]$/.test(node.tagName)) return;
    
    // Skip list elements (already processed)
    if (node.nodeType === Node.ELEMENT_NODE && (node.tagName === 'UL' || node.tagName === 'OL')) return;
    
    // Get text content
    let text = '';
    if (node.nodeType === Node.TEXT_NODE) {
      text = node.textContent.trim();
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      text = node.textContent.trim();
    }
    
    // Skip empty or nbsp-only
    if (!text || /^[\s\u00A0]*$/.test(text)) return;
    
    // Skip already captured
    if (capturedTexts.has(text)) return;
    
    // Check if substring of captured
    let isSubstring = false;
    for (const captured of capturedTexts) {
      if (captured.includes(text) && text.length <= captured.length) {
        isSubstring = true;
        break;
      }
    }
    if (isSubstring) return;
    
    // Determine if italic or bold (check both the element itself and its children)
    let isItalic = false;
    let isBold = false;
    
    if (node.nodeType === Node.ELEMENT_NODE) {
      isItalic = node.tagName === 'I' || node.tagName === 'EM' || node.querySelector('i, em');
      isBold = node.tagName === 'B' || node.tagName === 'STRONG' || node.querySelector('b, strong');
    }
    
    let textType = 'text';
    if (isBold && isItalic) textType = 'bolditalic';
    else if (isBold) textType = 'bold';
    else if (isItalic) textType = 'italic';
    
    console.log('Step 3b: Found top-level inline text (' + textType + '):', text.substring(0, 50));
    
    inlineElements.push({ 
      type: textType, 
      content: text, 
      element: node 
    });
    capturedTexts.add(text);
  });
  
  console.log('Step 3b complete: Found', inlineElements.length, 'inline elements');
  
  // Step 4: Combine all content and sort by DOM position
  const allContent = [...headingData, ...listItems, ...textBlocks, ...inlineElements];
  
  // Sort by document position
  allContent.sort((a, b) => {
    const position = a.element.compareDocumentPosition(b.element);
    if (position & Node.DOCUMENT_POSITION_FOLLOWING) return -1;
    if (position & Node.DOCUMENT_POSITION_PRECEDING) return 1;
    return 0;
  });
  
  // Step 5: Build final results (without element references)
  const finalResults = allContent.map(item => ({
    type: item.type,
    content: item.content,
    htmlContent: item.htmlContent || null,
    fontSize: item.fontSize || null,
    level: item.level || null
  }));
  
  // Deduplicate any remaining duplicates
  const deduped = [];
  finalResults.forEach((el, idx) => {
    // Keep breaks, but skip if immediately after a heading
    if (el.type === 'break') {
      const prev = deduped[deduped.length - 1];
      if (prev && (prev.type === 'h1' || prev.type === 'h2' || prev.type === 'h3')) {
        // Skip break immediately after heading
        return;
      }
      deduped.push(el);
      return;
    }
    if (!el.content || el.content.length === 0) return;
    const isDupe = deduped.some(d => d.content === el.content);
    if (!isDupe) deduped.push(el);
  });
  
  console.log('=== parseParagraphContent RESULT ===');
  console.log('Final deduped items:', deduped.length);
  deduped.forEach((item, i) => console.log(i, item.type, item.content?.substring(0, 40) || '(break)'));
  console.log('=== parseParagraphContent END ===');
  
  return { parsedContent: deduped, content: element.textContent.trim() };
};

// Parse HTML table - preserve rich content in cells
const parseTable = (table) => {
  const rows = [];
  const tableRows = table.querySelectorAll('tr');
  
  tableRows.forEach(tr => {
    const row = [];
    const cells = tr.querySelectorAll('td, th');
    
    cells.forEach(cell => {
      // Clone the cell to manipulate it
      const clone = cell.cloneNode(true);
      
      // Replace <br> tags with paragraph break markers
      clone.querySelectorAll('br').forEach(br => {
        br.replaceWith('\n\n');
      });
      
      // Replace divs with paragraph break + content (divs represent paragraphs)
      const divs = clone.querySelectorAll('div');
      divs.forEach(div => {
        const text = div.textContent.trim();
        if (text) {
          div.insertAdjacentText('beforebegin', '\n\n');
        }
      });
      
      // Get plain text content for the content field
      let cellContent = clone.textContent;
      cellContent = cellContent.replace(/\u00A0+/g, ' ');
      cellContent = cellContent.replace(/\n\s*\n/g, '\n\n');
      cellContent = cellContent.replace(/[ \t]+/g, ' ');
      cellContent = cellContent.split('\n').map(line => line.trim()).join('\n');
      cellContent = cellContent.trim();
      
      // Now get HTML content preserving inline formatting
      // Re-clone to get fresh HTML
      const htmlClone = cell.cloneNode(true);
      
      // Replace <br> tags with paragraph break marker
      htmlClone.querySelectorAll('br').forEach(br => {
        const marker = document.createTextNode('\n\n');
        br.replaceWith(marker);
      });
      
      // Process divs to add paragraph breaks
      htmlClone.querySelectorAll('div').forEach(div => {
        const text = div.textContent.trim();
        if (text) {
          div.insertAdjacentText('beforebegin', '\n\n');
        }
      });
      
      // Get innerHTML and clean it up, but preserve b, strong, i, em
      let htmlContent = htmlClone.innerHTML;
      
      // Remove span wrappers but keep content
      htmlContent = htmlContent.replace(/<span[^>]*>(.*?)<\/span>/gi, '$1');
      // Remove empty tags
      htmlContent = htmlContent.replace(/<(\w+)[^>]*>\s*<\/\1>/gi, '');
      // Remove div tags but keep content
      htmlContent = htmlContent.replace(/<\/?div[^>]*>/gi, '');
      // Convert nbsp
      htmlContent = htmlContent.replace(/&nbsp;/gi, ' ');
      // Remove links but keep text
      htmlContent = htmlContent.replace(/<a[^>]*>(.*?)<\/a>/gi, '$1');
      // Remove underline tags
      htmlContent = htmlContent.replace(/<\/?u>/gi, '');
      // Clean up whitespace
      htmlContent = htmlContent.replace(/\n\s*\n/g, '\n\n');
      htmlContent = htmlContent.replace(/[ \t]+/g, ' ');
      htmlContent = htmlContent.split('\n').map(line => line.trim()).join('\n');
      htmlContent = htmlContent.trim();
      
      // Check if there's any actual formatting
      const hasFormatting = /<(b|strong|i|em)[^>]*>/i.test(htmlContent);
      
      row.push({
        content: cellContent,
        htmlContent: hasFormatting ? htmlContent : null
      });
    });
    
    if (row.length > 0) rows.push(row);
  });
  
  return rows;
};

// Extract sections (H1-H3) for filtering - now includes content ranges
const extractSections = (content) => {
  const sections = [];
  let sectionId = 0;
  
  content.forEach((item, idx) => {
    if (item.type === 'divider' && item.title) {
      sections.push({ id: sectionId++, title: item.title, level: 1, contentIndex: idx, type: 'divider', selected: true });
    }
    else if (item.type === 'accordion') {
      sections.push({ id: sectionId++, title: item.title, level: 2, contentIndex: idx, type: 'accordion', selected: true });
    }
    else if (item.type === 'paragraph' && item.parsedContent) {
      item.parsedContent.forEach((el, elIdx) => {
        if (['h1', 'h2', 'h3'].includes(el.type)) {
          const level = parseInt(el.type.charAt(1));
          sections.push({ id: sectionId++, title: el.content, level, contentIndex: idx, type: 'paragraph-header', elementIndex: elIdx, selected: true });
        }
      });
    }
  });
  
  return sections;
};

// Generate HTML for Word
const generateHtmlForWord = (content) => {
  const headerStyles = {
    h1: "mso-style-name:'Heading 1'; mso-outline-level:1; font-size:28px; color:#1a202c; margin:24px 0 12px 0; font-weight:bold;",
    h2: "mso-style-name:'Heading 2'; mso-outline-level:2; font-size:22px; color:#2d3748; margin:20px 0 10px 0; font-weight:bold;",
    h3: "mso-style-name:'Heading 3'; mso-outline-level:3; font-size:16px; color:#4a5568; margin:16px 0 8px 0; font-weight:bold;",
  };
  
  const tableStyle = `style="width:100%; table-layout:fixed; border-collapse:collapse; margin:12px 0;"`;
  const cellStyle = (isHeader) => `style="border:1px solid #333; padding:8px; word-wrap:break-word;${isHeader ? ' font-weight:bold; background-color:#f0f0f0;' : ''}"`;
  const spacerParagraph = '<p style="mso-style-name:\'Normal\'; margin:0; line-height:100%;"><span style="font-size:11pt;">&nbsp;</span></p>';
  
  // Check if parsedContent ends with a break
  const endsWithBreak = (parsedContent) => {
    if (!parsedContent || parsedContent.length === 0) return false;
    return parsedContent[parsedContent.length - 1].type === 'break';
  };
  
  // Check if parsedContent starts with a heading (which has its own margin)
  const startsWithHeading = (parsedContent) => {
    if (!parsedContent || parsedContent.length === 0) return false;
    const firstType = parsedContent[0].type;
    return ['h1', 'h2', 'h3'].includes(firstType);
  };
  
  const renderParsedContent = (parsedContent) => {
    let html = '';
    let currentListLevel = 0;
    
    // Helper to check if a type is a list item
    const isListItem = (type) => type === 'listitem';
    
    parsedContent.forEach((el, idx) => {
      const elLevel = el.level || 1;
      
      // Handle list level changes
      if (isListItem(el.type)) {
        // Open new lists to reach the target level
        while (currentListLevel < elLevel) {
          const indent = currentListLevel * 20;
          html += `<ul style="margin:8px 0 8px ${20 + indent}px;">`;
          currentListLevel++;
        }
        // Close lists to reach the target level
        while (currentListLevel > elLevel) {
          html += '</ul>';
          currentListLevel--;
        }
        
        // Render the list item with htmlContent if available
        const liContent = el.htmlContent || el.content;
        html += `<li style="margin:4px 0;">${liContent}</li>`;
      } else {
        // Close all open lists when leaving list items
        while (currentListLevel > 0) {
          html += '</ul>';
          currentListLevel--;
        }
        
        // Build font size style if present
        const fontStyle = el.fontSize ? `font-size:${el.fontSize}pt;` : '';
        
        if (el.type === 'h1') html += `<h1 style="${headerStyles.h1}">${el.content}</h1>`;
        else if (el.type === 'h2') html += `<h2 style="${headerStyles.h2}">${el.content}</h2>`;
        else if (el.type === 'h3') html += `<h3 style="${headerStyles.h3}">${el.content}</h3>`;
        else if (el.type === 'break') {
          html += spacerParagraph;
        }
        else if (el.type === 'bold') html += `<p style="margin:0 0 0 0;${fontStyle}"><strong>${el.content}</strong></p>`;
        else if (el.type === 'italic') html += `<p style="margin:0 0 0 0;${fontStyle}"><em>${el.content}</em></p>`;
        else if (el.type === 'bolditalic') html += `<p style="margin:0 0 0 0;${fontStyle}"><strong><em>${el.content}</em></strong></p>`;
        else {
          const displayContent = el.htmlContent || el.content;
          html += `<p style="margin:0 0 0 0;${fontStyle}">${displayContent}</p>`;
        }
      }
    });
    
    // Close any remaining open lists
    while (currentListLevel > 0) {
      html += '</ul>';
      currentListLevel--;
    }
    
    return html;
  };
  
  const renderTable = (tableData) => {
    if (!tableData?.length) return '';
    let html = `<table ${tableStyle}>`;
    tableData.forEach((row, rIdx) => {
      html += '<tr>';
      row.forEach(cell => {
        // Cell can be an object with content/htmlContent or a plain string (backward compat)
        const cellText = typeof cell === 'object' ? cell.content : cell;
        const cellHtmlContent = typeof cell === 'object' ? cell.htmlContent : null;
        
        // Use htmlContent if available, otherwise plain content
        const sourceContent = cellHtmlContent || cellText;
        
        // Convert paragraph breaks (\n\n) to HTML paragraphs within cells
        const cellHtml = sourceContent
          .split(/\n\n+/)
          .map(para => para.trim())
          .filter(para => para.length > 0)
          .map(para => `<p style="margin:4px 0;">${para}</p>`)
          .join('');
        html += `<td ${cellStyle(rIdx === 0)}>${cellHtml || sourceContent}</td>`;
      });
      html += '</tr>';
    });
    return html + '</table>';
  };
  
  const renderContent = (item) => {
    let html = '';
    switch (item.type) {
      case 'paragraph':
        if (item.parsedContent?.length) {
          html += renderParsedContent(item.parsedContent);
        } else if (item.content) {
          html += `<p style="margin:6px 0;">${item.content}</p>`;
        }
        break;
      case 'table':
        html += renderTable(item.table);
        break;
      case 'divider':
        if (item.title) {
          html += `<h1 style="${headerStyles.h1}">${item.title}</h1>`;
        }
        break;
      case 'accordion':
        html += `<h2 style="${headerStyles.h2}">${item.title}</h2>`;
        item.children?.forEach(child => { html += renderContent(child); });
        break;
    }
    return html;
  };
  
  // Determine if we need a spacer between two content items
  const needsSpacerBetween = (prevItem, currentItem) => {
    if (!prevItem) return false;
    
    // Dividers/accordions have their own spacing via headings
    if (currentItem.type === 'divider' || currentItem.type === 'accordion') return false;
    if (prevItem.type === 'divider' || prevItem.type === 'accordion') return false;
    
    // If current item is a paragraph starting with a heading, it has its own margin
    if (currentItem.type === 'paragraph' && startsWithHeading(currentItem.parsedContent)) return false;
    
    // If previous item was a paragraph that ended with a break, no need for spacer
    if (prevItem.type === 'paragraph' && endsWithBreak(prevItem.parsedContent)) return false;
    
    // Between different block types (e.g., paragraph -> table, table -> paragraph), add spacer
    if (prevItem.type !== currentItem.type) return true;
    
    // Between two tables, add spacer
    if (prevItem.type === 'table' && currentItem.type === 'table') return true;
    
    return false;
  };
  
  let fullHtml = `<html><head><style>body{font-family:Calibri,Arial,sans-serif;font-size:11pt;max-width:100%;}</style></head><body>`;
  
  content.forEach((item, idx) => {
    const prevItem = idx > 0 ? content[idx - 1] : null;
    
    // Add spacer between blocks if needed
    if (needsSpacerBetween(prevItem, item)) {
      fullHtml += spacerParagraph;
    }
    
    fullHtml += renderContent(item);
  });
  
  fullHtml += '</body></html>';
  return fullHtml;
};

export default function SeismicConverter() {
  const [inputText, setInputText] = useState('');
  const [docContent, setDocContent] = useState(null);
  const [sections, setSections] = useState([]);
  const [showPreview, setShowPreview] = useState(false);
  const [activeTab, setActiveTab] = useState('input');
  const [status, setStatus] = useState({ type: '', message: '' });

  const parseInput = useCallback(() => {
    try {
      const trimmed = inputText.trim();
      if (!trimmed) {
        setStatus({ type: 'error', message: 'Please paste Seismic HTML content' });
        return null;
      }
      
      const content = parseSeismicHtml(trimmed);
      
      if (content.length === 0) {
        setStatus({ type: 'error', message: 'No content found. Make sure you copied the Seismic article HTML.' });
        return null;
      }
      
      setDocContent(content);
      const extractedSections = extractSections(content);
      setSections(extractedSections);
      
      const tableCount = content.filter(c => c.type === 'table').length;
      setStatus({ type: 'success', message: `Parsed ${content.length} blocks (${tableCount} tables)` });
      return content;
    } catch (err) {
      setStatus({ type: 'error', message: 'Error parsing content: ' + err.message });
      setDocContent(null);
      setSections([]);
      return null;
    }
  }, [inputText]);

  const toggleSection = useCallback((sectionId) => {
    setSections(prev => {
      const updated = [...prev];
      const idx = updated.findIndex(s => s.id === sectionId);
      if (idx === -1) return prev;
      
      const section = updated[idx];
      const newSelected = !section.selected;
      updated[idx] = { ...section, selected: newSelected };
      
      if (!newSelected) {
        for (let i = idx + 1; i < updated.length; i++) {
          if (updated[i].level > section.level) {
            updated[i] = { ...updated[i], selected: false };
          } else break;
        }
      }
      
      if (newSelected) {
        for (let i = idx - 1; i >= 0; i--) {
          if (updated[i].level < section.level) {
            updated[i] = { ...updated[i], selected: true };
            if (updated[i].level === 1) break;
          }
        }
      }
      
      return updated;
    });
  }, []);

  const selectAll = useCallback(() => setSections(prev => prev.map(s => ({ ...s, selected: true }))), []);
  const deselectAll = useCallback(() => setSections(prev => prev.map(s => ({ ...s, selected: false }))), []);

  // FIXED: getFilteredContent now properly includes content between sections
  const getFilteredContent = useCallback(() => {
    if (!docContent) return null;
    
    // If no sections defined, return all content
    if (sections.length === 0) return docContent;
    
    // Build a map of which content indices belong to which section
    // Each content item belongs to the most recent section before it
    const sectionsByContentIndex = new Map();
    sections.forEach(s => {
      sectionsByContentIndex.set(s.contentIndex, s);
    });
    
    // For each content item, find its "owning" section (the nearest section at or before this index)
    const getOwningSection = (contentIdx) => {
      // Check if this content item IS a section
      if (sectionsByContentIndex.has(contentIdx)) {
        return sectionsByContentIndex.get(contentIdx);
      }
      
      // Find the nearest section before this index
      let owningSection = null;
      for (let i = contentIdx - 1; i >= 0; i--) {
        if (sectionsByContentIndex.has(i)) {
          owningSection = sectionsByContentIndex.get(i);
          break;
        }
      }
      return owningSection;
    };
    
    // Filter content: include if it's a selected section OR belongs to a selected section
    return docContent.filter((item, idx) => {
      // If this item is itself a section header, check if it's selected
      const asSection = sectionsByContentIndex.get(idx);
      if (asSection) {
        return asSection.selected;
      }
      
      // Otherwise, include if its owning section is selected
      const owner = getOwningSection(idx);
      
      // If no owning section (content before first section), include it
      if (!owner) return true;
      
      return owner.selected;
    });
  }, [docContent, sections]);

  const handlePreview = useCallback(() => {
    const content = docContent || parseInput();
    if (content) {
      setShowPreview(true);
      setActiveTab('preview');
    }
  }, [docContent, parseInput]);

  const handleClear = useCallback(() => {
    setInputText('');
    setDocContent(null);
    setSections([]);
    setShowPreview(false);
    setActiveTab('input');
    setStatus({ type: '', message: '' });
  }, []);

  const handleCopyForWord = useCallback(() => {
    let content = getFilteredContent();
    if (!content) {
      content = parseInput();
      if (!content) return;
    }
    
    const html = generateHtmlForWord(content);
    const blob = new Blob([html], { type: 'text/html' });
    const clipboardItem = new ClipboardItem({ 'text/html': blob });
    
    navigator.clipboard.write([clipboardItem]).then(() => {
      setStatus({ type: 'success', message: 'Copied! Paste directly into Word (Ctrl+V / Cmd+V)' });
    }).catch(() => {
      navigator.clipboard.writeText(html).then(() => {
        setStatus({ type: 'success', message: 'Copied as HTML. Use Paste Special > HTML in Word.' });
      }).catch(err => {
        setStatus({ type: 'error', message: 'Failed to copy: ' + err.message });
      });
    });
  }, [getFilteredContent, parseInput]);

  const filteredContent = useMemo(() => getFilteredContent(), [getFilteredContent]);

  const renderPreview = () => {
    const content = filteredContent || docContent;
    if (!content || !showPreview) return null;
    
    return (
      <div className="p-4 bg-white rounded border max-h-96 overflow-y-auto">
        <h3 className="font-bold text-lg mb-3 border-b pb-2">Document Preview ({content.length} items)</h3>
        {content.map((item, idx) => (
          <div key={idx} className="mb-3">
            {item.type === 'divider' && item.title && (
              <div className="mt-4">
                <h1 className="text-2xl font-bold text-gray-900 my-3">{item.title}</h1>
              </div>
            )}
            {item.type === 'paragraph' && (
              <div className="text-sm">
                {item.parsedContent?.map((el, elIdx) => {
                  if (el.type === 'h1') return <h1 key={elIdx} className="text-2xl font-bold text-gray-900 my-3">{el.content}</h1>;
                  if (el.type === 'h2') return <h2 key={elIdx} className="text-xl font-bold text-gray-800 my-2">{el.content}</h2>;
                  if (el.type === 'h3') return <h3 key={elIdx} className="text-lg font-bold text-gray-700 my-2">{el.content}</h3>;
                  if (el.type === 'listitem') {
                    const indent = ((el.level || 1) - 1) * 16;
                    if (el.htmlContent) {
                      return <li key={elIdx} className="ml-4" style={{ marginLeft: `${16 + indent}px` }} dangerouslySetInnerHTML={{ __html: el.htmlContent }} />;
                    }
                    return <li key={elIdx} className="ml-4" style={{ marginLeft: `${16 + indent}px` }}>{el.content}</li>;
                  }
                  if (el.type === 'bold') return <p key={elIdx} className="font-bold">{el.content}</p>;
                  if (el.type === 'italic') return <p key={elIdx} className="italic">{el.content}</p>;
                  if (el.type === 'bolditalic') return <p key={elIdx} className="font-bold italic">{el.content}</p>;
                  if (el.type === 'break') return <div key={elIdx} className="h-4"></div>;
                  // Use htmlContent if available (preserves inline bold/italic)
                  if (el.htmlContent) {
                    return <p key={elIdx} className="text-gray-700" dangerouslySetInnerHTML={{ __html: el.htmlContent }} />;
                  }
                  return <p key={elIdx} className="text-gray-700">{el.content}</p>;
                })}
              </div>
            )}
            {item.type === 'table' && item.table && (
              <div className="my-4 overflow-x-auto">
                <table className="text-xs border-collapse w-full">
                  <tbody>
                    {item.table.map((row, rIdx) => (
                      <tr key={rIdx}>
                        {row.map((cell, cIdx) => {
                          // Cell can be an object with content/htmlContent or a plain string
                          const cellText = typeof cell === 'object' ? cell.content : cell;
                          const cellHtmlContent = typeof cell === 'object' ? cell.htmlContent : null;
                          const sourceContent = cellHtmlContent || cellText;
                          
                          return (
                            <td key={cIdx} className={`border border-gray-300 p-2 ${rIdx === 0 ? 'font-bold bg-gray-100' : ''}`}>
                              {sourceContent.split(/\n\n+/).map((para, pIdx) => (
                                cellHtmlContent 
                                  ? <p key={pIdx} className={pIdx > 0 ? 'mt-2' : ''} dangerouslySetInnerHTML={{ __html: para }} />
                                  : <p key={pIdx} className={pIdx > 0 ? 'mt-2' : ''}>{para}</p>
                              ))}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
            {item.type === 'accordion' && (
              <div className="my-3 ml-2 relative pl-5">
                <div className="absolute left-0 top-0 bottom-0 w-1 bg-green-400 rounded-full"></div>
                <h2 className="text-xl font-bold text-gray-900 mb-2">{item.title}</h2>
                {item.children?.map((child, cIdx) => (
                  <div key={cIdx} className="mt-2 text-sm">
                    {child.type === 'paragraph' && child.parsedContent && (
                      <div>
                        {child.parsedContent.map((el, elIdx) => {
                          if (el.type === 'h1') return <h1 key={elIdx} className="text-xl font-bold text-gray-900 my-2">{el.content}</h1>;
                          if (el.type === 'h2') return <h2 key={elIdx} className="text-lg font-bold text-gray-800 my-2">{el.content}</h2>;
                          if (el.type === 'h3') return <h3 key={elIdx} className="text-base font-bold text-gray-700 my-1">{el.content}</h3>;
                          if (el.type === 'listitem') {
                            const indent = ((el.level || 1) - 1) * 16;
                            if (el.htmlContent) {
                              return <li key={elIdx} className="ml-4" style={{ marginLeft: `${16 + indent}px` }} dangerouslySetInnerHTML={{ __html: el.htmlContent }} />;
                            }
                            return <li key={elIdx} className="ml-4" style={{ marginLeft: `${16 + indent}px` }}>{el.content}</li>;
                          }
                          if (el.type === 'break') return <div key={elIdx} className="h-3"></div>;
                          if (el.htmlContent) {
                            return <p key={elIdx} className="text-gray-600" dangerouslySetInnerHTML={{ __html: el.htmlContent }} />;
                          }
                          return <p key={elIdx} className="text-gray-600">{el.content}</p>;
                        })}
                      </div>
                    )}
                    {child.type === 'paragraph' && !child.parsedContent && child.content && (
                      <p className="text-gray-600">{child.content}</p>
                    )}
                    {child.type === 'table' && child.table && (
                      <table className="text-xs border-collapse w-full mt-2">
                        <tbody>
                          {child.table.map((row, rIdx) => (
                            <tr key={rIdx}>
                              {row.map((cell, cIdx2) => {
                                const cellText = typeof cell === 'object' ? cell.content : cell;
                                const cellHtmlContent = typeof cell === 'object' ? cell.htmlContent : null;
                                const sourceContent = cellHtmlContent || cellText;
                                
                                return (
                                  <td key={cIdx2} className={`border border-gray-300 p-1 ${rIdx === 0 ? 'font-bold bg-gray-50' : ''}`}>
                                    {sourceContent.split(/\n\n+/).map((para, pIdx) => (
                                      cellHtmlContent
                                        ? <p key={pIdx} className={pIdx > 0 ? 'mt-1' : ''} dangerouslySetInnerHTML={{ __html: para }} />
                                        : <p key={pIdx} className={pIdx > 0 ? 'mt-1' : ''}>{para}</p>
                                    ))}
                                  </td>
                                );
                              })}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </div>
                ))}
              </div>
            )}
          </div>
        ))}
      </div>
    );
  };

  const renderSectionsTab = () => {
    if (sections.length === 0) {
      return (
        <div className="p-4 text-gray-500 text-center">
          No sections found. Paste HTML and click Preview first.
        </div>
      );
    }
    
    return (
      <div className="p-4">
        <div className="flex gap-2 mb-4">
          <button onClick={selectAll} className="flex items-center gap-1 px-3 py-1 text-sm bg-gray-100 hover:bg-gray-200 rounded">
            <CheckSquare className="w-4 h-4" /> Select All
          </button>
          <button onClick={deselectAll} className="flex items-center gap-1 px-3 py-1 text-sm bg-gray-100 hover:bg-gray-200 rounded">
            <Square className="w-4 h-4" /> Deselect All
          </button>
        </div>
        <div className="border rounded max-h-80 overflow-y-auto">
          {sections.map((section) => (
            <div
              key={section.id}
              className={`flex items-center gap-2 p-2 hover:bg-gray-50 border-b last:border-b-0 cursor-pointer ${!section.selected ? 'opacity-50' : ''}`}
              style={{ paddingLeft: `${(section.level - 1) * 20 + 8}px` }}
              onClick={() => toggleSection(section.id)}
            >
              <input type="checkbox" checked={section.selected} onChange={() => toggleSection(section.id)} className="w-4 h-4 rounded" onClick={(e) => e.stopPropagation()} />
              <span className={`text-xs px-1.5 py-0.5 rounded ${section.level === 1 ? 'bg-blue-100 text-blue-700' : section.level === 2 ? 'bg-green-100 text-green-700' : 'bg-gray-100 text-gray-600'}`}>
                H{section.level}
              </span>
              <span className={`text-sm ${section.level === 1 ? 'font-bold' : section.level === 2 ? 'font-semibold' : ''}`}>
                {section.title}
              </span>
            </div>
          ))}
        </div>
        <div className="mt-3 text-sm text-gray-500">
          {sections.filter(s => s.selected).length} of {sections.length} sections selected
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-gray-100 p-4">
      <div className="max-w-4xl mx-auto">
        <div className="bg-white rounded-lg shadow-lg p-5">
          <div className="flex items-center gap-3 mb-4">
            <FileText className="w-7 h-7 text-blue-600" />
            <h1 className="text-xl font-bold text-gray-800">Seismic HTML to DOCX Converter</h1>
          </div>
          
          <div className="flex border-b mb-4">
            <button onClick={() => setActiveTab('input')} className={`px-4 py-2 text-sm font-medium border-b-2 transition ${activeTab === 'input' ? 'border-blue-600 text-blue-600' : 'border-transparent text-gray-500 hover:text-gray-700'}`}>
              Input HTML
            </button>
            <button onClick={() => { if (docContent) setActiveTab('sections'); }} className={`px-4 py-2 text-sm font-medium border-b-2 transition flex items-center gap-1 ${activeTab === 'sections' ? 'border-blue-600 text-blue-600' : 'border-transparent text-gray-500 hover:text-gray-700'} ${!docContent ? 'opacity-50 cursor-not-allowed' : ''}`}>
              <List className="w-4 h-4" /> Sections
            </button>
            <button onClick={() => { if (docContent) { setShowPreview(true); setActiveTab('preview'); } }} className={`px-4 py-2 text-sm font-medium border-b-2 transition flex items-center gap-1 ${activeTab === 'preview' ? 'border-blue-600 text-blue-600' : 'border-transparent text-gray-500 hover:text-gray-700'} ${!docContent ? 'opacity-50 cursor-not-allowed' : ''}`}>
              <Eye className="w-4 h-4" /> Preview
            </button>
          </div>

          {activeTab === 'input' && (
            <textarea
              value={inputText}
              onChange={(e) => setInputText(e.target.value)}
              placeholder="Paste your Seismic HTML here (use the bookmarklet to copy it)..."
              className="w-full h-48 p-3 border border-gray-300 rounded-lg font-mono text-sm resize-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
            />
          )}
          
          {activeTab === 'sections' && renderSectionsTab()}
          {activeTab === 'preview' && renderPreview()}
          
          <div className="flex gap-3 mt-4">
            <button onClick={handlePreview} disabled={!inputText.trim()} className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition">
              <Eye className="w-4 h-4" /> Preview
            </button>
            <button onClick={handleCopyForWord} disabled={!inputText.trim()} className="flex items-center gap-2 px-4 py-2 bg-orange-500 text-white rounded hover:bg-orange-600 disabled:bg-gray-300 disabled:cursor-not-allowed transition">
              <ClipboardPaste className="w-4 h-4" /> Copy for Word
            </button>
            <button onClick={handleClear} className="flex items-center gap-2 px-4 py-2 bg-red-500 text-white rounded hover:bg-red-600 transition">
              <Trash2 className="w-4 h-4" /> Clear
            </button>
          </div>

          {status.message && (
            <div className={`flex items-center gap-2 p-3 rounded mt-4 ${status.type === 'error' ? 'bg-red-100 text-red-700' : 'bg-green-100 text-green-700'}`}>
              {status.type === 'error' ? <AlertCircle className="w-5 h-5" /> : <CheckCircle className="w-5 h-5" />}
              <span className="text-sm">{status.message}</span>
            </div>
          )}

          <div className="mt-4 p-3 bg-amber-50 border border-amber-200 rounded text-sm text-amber-800">
            <strong>Tip:</strong> Use the bookmarklet on a Seismic page to copy the HTML, then paste it here.
          </div>
        </div>
      </div>
    </div>
  );
}
