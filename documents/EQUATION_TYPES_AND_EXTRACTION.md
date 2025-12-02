# Equation Types and Extraction Methods

## Overview

Word documents store equations (OMML - Office Math Markup Language) in different XML locations depending on how they were inserted. Understanding these locations is crucial for complete equation extraction.

---

## Equation Storage Locations

### 1. Main Body Equations
- **Location**: Directly in `document.xml` body
- **Path**: `w:document > w:body > w:p > m:oMath`
- **Access**: Easy - both ZIP and COM methods can extract

### 2. Shape/Textbox Equations (Modern - mc:Choice)
- **Location**: Inside `mc:AlternateContent > mc:Choice` blocks
- **Path**: `mc:Choice > ... > wps:txbx > w:txbxContent > w:p > m:oMath`
- **Access**: ZIP can extract directly; COM uses `Shape.TextFrame.TextRange.OMaths`
- **Note**: These are the "real" equations for Word 2010+

### 3. Shape/Textbox Equations (Legacy - mc:Fallback)
- **Location**: Inside `mc:AlternateContent > mc:Fallback` blocks
- **Path**: `mc:Fallback > w:pict > v:oval > v:textbox > w:txbxContent > w:p > m:oMath`
- **Access**: ZIP can extract; COM does NOT access these (they're duplicates)
- **Note**: These are DUPLICATE copies for older Word versions

---

## Key Insight: Duplicate Equations

When a document contains equations inside shapes/textboxes, Word stores them **TWICE**:

1. **mc:Choice** - Modern version (for Word 2010+)
2. **mc:Fallback** - Legacy version (for older Word versions)

### Example from Test Document:

```
Total m:oMath elements in XML: 144

Breakdown:
  Main body:     70  (unique)
  mc:Choice:     37  (unique - in shapes)
  mc:Fallback:   37  (DUPLICATES of mc:Choice)

ACTUAL UNIQUE EQUATIONS: 107
FALLBACK DUPLICATES:     37
```

---

## Extraction Methods Comparison

### ZIP Method (Pure Python)

**Advantages:**
- Cross-platform (Linux, Windows, Mac)
- No Word installation required
- Low cost (~$50/month server)
- Can access ALL equation locations including fallbacks

**Process:**
1. Extract .docx (it's a ZIP file)
2. Parse `word/document.xml` with lxml
3. Find equations in ALL locations using XPath
4. Convert OMML → LaTeX
5. Replace equations with marked text
6. Repackage .docx

**XPath for all equations:**
```python
# All equations (including duplicates)
all_eqs = root.xpath('//m:oMath', namespaces=ns)

# Only unique equations (exclude fallbacks)
unique_eqs = root.xpath('//m:oMath[not(ancestor::mc:Fallback)]', namespaces=ns)
```

### Word COM Method (Windows only)

**Advantages:**
- Native Word API access
- Better handling of complex documents
- Can access equations through multiple methods

**Process:**
1. Open document via COM
2. Access `doc.OMaths` for main body
3. Access `Shape.TextFrame.TextRange.OMaths` for shapes
4. Replace equations using Range operations

**Note:** COM only accesses mc:Choice equations, not mc:Fallback. This is correct behavior as fallbacks are duplicates.

---

## Conversion Results

### Test Document 1: الدالة واحد لواحد
- **XML Total**: 144 m:oMath elements
- **Unique Equations**: 107 (70 main + 37 in shapes)
- **Fallback Duplicates**: 37
- **ZIP Conversion**: 144/144 (100%) - converts all including fallbacks
- **COM Conversion**: 107/107 (100%) - converts unique equations only

### Test Document 2: التشابه
- **XML Total**: 89 m:oMath elements
- **Unique Equations**: 89 (all in main body)
- **Fallback Duplicates**: 0
- **Both Methods**: 89/89 (100%)

---

## Recommended Approach

### For Maximum Compatibility (ZIP Method):
Convert ALL equations including mc:Fallback duplicates. This ensures:
- Modern Word versions see converted equations in mc:Choice
- Older Word versions see converted equations in mc:Fallback

### For Efficiency (COM Method):
Convert only unique equations (107 instead of 144). The fallback versions will remain as OMML but won't be displayed in modern Word.

---

## Code Examples

### ZIP: Find All Equation Locations
```python
def get_equation_location(eq, namespaces):
    is_in_choice = bool(eq.xpath('ancestor::mc:Choice', namespaces=namespaces))
    is_in_fallback = bool(eq.xpath('ancestor::mc:Fallback', namespaces=namespaces))

    if is_in_choice:
        return 'mc_choice'      # Modern shape equation
    elif is_in_fallback:
        return 'mc_fallback'    # Legacy duplicate
    else:
        return 'main_body'      # Regular equation
```

### COM: Access Shape Equations
```python
# Main document equations
for i in range(1, doc.OMaths.Count + 1):
    eq = doc.OMaths.Item(i)
    # Process equation

# Shape/textbox equations
for i in range(1, doc.Shapes.Count + 1):
    shape = doc.Shapes.Item(i)
    if shape.TextFrame.HasText:
        text_range = shape.TextFrame.TextRange
        for j in range(1, text_range.OMaths.Count + 1):
            eq = text_range.OMaths.Item(j)
            # Process equation
```

---

## Summary

| Location | ZIP Access | COM Access | Is Duplicate? |
|----------|------------|------------|---------------|
| Main body | ✅ Yes | ✅ Yes | No |
| mc:Choice (shapes) | ✅ Yes | ✅ Yes (via Shape.TextFrame) | No |
| mc:Fallback (shapes) | ✅ Yes | ❌ No (not needed) | Yes |

**Both methods can achieve 100% conversion of all unique equations!**
