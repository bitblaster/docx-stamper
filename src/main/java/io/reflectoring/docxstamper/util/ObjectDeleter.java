package io.reflectoring.docxstamper.util;

import jakarta.xml.bind.JAXBElement;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
import io.reflectoring.docxstamper.api.coordinates.ParagraphCoordinates;
import io.reflectoring.docxstamper.api.coordinates.TableCoordinates;
import io.reflectoring.docxstamper.api.coordinates.TableRowCoordinates;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ObjectDeleter {

	private WordprocessingMLPackage document;

	private List<Integer> deletedObjectsIndexes = new ArrayList<>(10);

	private Map<ContentAccessor, Integer> deletedObjectsPerParent = new HashMap<>();

	public ObjectDeleter(WordprocessingMLPackage document) {
		this.document = document;
	}

	public void deleteParagraph(ParagraphCoordinates paragraphCoordinates) {
		if (paragraphCoordinates.getParentTableCellCoordinates() == null) {
			// global paragraph
			int indexToDelete = getOffset(paragraphCoordinates.getIndex());
			document.getMainDocumentPart().getContent().remove(indexToDelete);
			deletedObjectsIndexes.add(paragraphCoordinates.getIndex());
		} else {
			// paragraph within a table cell
			Tc parentCell = paragraphCoordinates.getParentTableCellCoordinates().getCell();
			deleteFromCell(parentCell, paragraphCoordinates.getIndex());
		}
	}

	/**
	 * Get new index of element to be deleted, taking into account previously
	 * deleted elements
	 *
	 * @param initialIndex
	 *            initial index of the element to be deleted
	 * @return the index of the item to be removed
	 */
	private int getOffset(final int initialIndex) {
		int newIndex = initialIndex;
		for (Integer deletedIndex : this.deletedObjectsIndexes) {
			if (initialIndex > deletedIndex) {
				newIndex--;
			}
		}
		return newIndex;
	}

	private void deleteFromCell(Tc cell, int index) {
		Integer objectsDeletedFromParent = deletedObjectsPerParent.get(cell);
		if (objectsDeletedFromParent == null) {
			objectsDeletedFromParent = 0;
		}
		index -= objectsDeletedFromParent;
		cell.getContent().remove(index);
		if (!TableCellUtil.hasAtLeastOneParagraphOrTable(cell)) {
			TableCellUtil.addEmptyParagraph(cell);
		}
		deletedObjectsPerParent.put(cell, objectsDeletedFromParent + 1);
		// TODO: find out why border lines are removed in some cells after having
		// deleted a paragraph
	}

	public void deleteTable(TableCoordinates tableCoordinates) {
		if (tableCoordinates.getParentTableCellCoordinates() == null) {
			// global table
			int indexToDelete = getOffset(tableCoordinates.getIndex());
			document.getMainDocumentPart().getContent().remove(indexToDelete);
			deletedObjectsIndexes.add(tableCoordinates.getIndex());
		} else {
			// nested table within an table cell
			Tc parentCell = tableCoordinates.getParentTableCellCoordinates().getCell();
			deleteFromCell(parentCell, tableCoordinates.getIndex());
		}
	}

	public void deleteTableRow(TableRowCoordinates tableRowCoordinates) {
		Tbl table = tableRowCoordinates.getParentTableCoordinates().getTable();
		Tr row = tableRowCoordinates.getRow();
		// Before actually deleting the row, check if some cell is starting a vertical merge
		// If so, deleting the row will also remove the merged content,
		// so instead we "hide" the row forcing its height to zero
		for (Object rawCol : row.getContent()) {
			if (rawCol instanceof JAXBElement<?> jaxbCol && jaxbCol.getValue() instanceof Tc col) {
				TcPr tcPr = col.getTcPr();
				if (tcPr.getVMerge() != null && "restart".equals(tcPr.getVMerge().getVal())) {
					// Search for an existing height definition on the row
					List<JAXBElement<?>> trStyles = row.getTrPr().getCnfStyleOrDivIdOrGridBefore();
					for (JAXBElement<?> trStyle : trStyles) {
						if (trStyle.getValue() instanceof CTHeight trHeight) {
							trHeight.setVal(BigInteger.ZERO);
							trHeight.setHRule(STHeightRule.EXACT);
							return;
						}
					}

					// if not found, create a new height definition
					ObjectFactory wmlObjectFactory = new ObjectFactory();
					CTHeight trHeight = wmlObjectFactory.createCTHeight();
					trHeight.setVal(BigInteger.ZERO);
					trHeight.setHRule(STHeightRule.EXACT);
					JAXBElement<org.docx4j.wml.CTHeight> heightWrapped = wmlObjectFactory.createCTTrPrBaseTrHeight(trHeight);
					trStyles.add(heightWrapped);
					return;
				}
			}
		}

		int index = tableRowCoordinates.getIndex();
		Integer objectsDeletedFromTable = deletedObjectsPerParent.get(table);
		if (objectsDeletedFromTable == null) {
			objectsDeletedFromTable = 0;
		}
		index -= objectsDeletedFromTable;
		table.getContent().remove(index);
		deletedObjectsPerParent.put(table, objectsDeletedFromTable + 1);
	}

}
