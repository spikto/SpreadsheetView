//
//  ViewController.swift
//  Spreadsheet
//
//  Created by Kishikawa Katsumi on 2017/06/01.
//  Copyright © 2017 Kishikawa Katsumi. All rights reserved.
//

import UIKit
import SpreadsheetView

public protocol SpreadsheetDelegate {
    func spreadsheet(_ spreadsheet: Spreadsheet, performCellAction cellRange: CellRange, intersection: CellRange?)
    func spreadsheet(_ spreadsheet: Spreadsheet, textShouldBeginEditingAt indexPath: IndexPath) -> Bool
    func spreadsheet(_ spreadsheet: Spreadsheet, textDidBeginEditingAt indexPath: IndexPath)
}

extension SpreadsheetDelegate {
    func spreadsheet(_ spreadsheet: Spreadsheet, textShouldBeginEditingAt indexPath: IndexPath) -> Bool { return true }
    func spreadsheet(_ spreadsheet: Spreadsheet, textDidBeginEditingAt indexPath: IndexPath) {}
}

class SpreadsheetDataSource: SpreadsheetViewDataSource {
    var numberOfColumns = 255
    var numberOfRows = 65535
    var mergedCells = [CellRange]()
    let mergedCellStore = MergedCellStore()

    var data = [IndexPath: String]()

    func numberOfColumns(in spreadsheetView: SpreadsheetView) -> Int {
        return numberOfColumns
    }

    func numberOfRows(in spreadsheetView: SpreadsheetView) -> Int {
        return numberOfRows
    }

    func spreadsheetView(_ spreadsheetView: SpreadsheetView, widthForColumn column: Int) -> CGFloat {
        return 120
    }

    func spreadsheetView(_ spreadsheetView: SpreadsheetView, heightForRow row: Int) -> CGFloat {
        return 30
    }

    func spreadsheetView(_ spreadsheetView: SpreadsheetView, cellForItemAt indexPath: IndexPath) -> Cell? {
        if let text = data[indexPath] {
            let cell = spreadsheetView.dequeueReusableCell(withReuseIdentifier: String(describing: TextCell.self), for: indexPath) as! TextCell
            cell.text = text
            return cell
        }
        return nil
    }

    func mergedCells(in spreadsheetView: SpreadsheetView) -> [CellRange] {
        return mergedCells
    }

    func mergeCells(cellRange: CellRange) {
        for indexPath in cellRange {
            if let existingMergedCell = mergedCellStore[indexPath] {
                if existingMergedCell.contains(cellRange) {
                    continue
                }
                if cellRange.contains(existingMergedCell) {
                    unmergeCell(cellRange: existingMergedCell)
                } else {
                    fatalError("cannot merge cells in a range that overlap existing merged cells")
                }
            }
            mergedCellStore[indexPath] = cellRange
        }
        mergedCells.append(cellRange)
    }

    func unmergeCell(cellRange: CellRange) {
        for indexPath in cellRange {
            if let range = mergedCellStore[indexPath] {
                if let index = mergedCells.index(of: range) {
                    mergedCells.remove(at: index)
                }
                mergedCellStore[indexPath] = nil
            }
        }
    }
}

class SelectionView: UIView {
    let leftCornerHandle = UIControl()
    let rightCornerHandle = UIControl()
    let borderLayer = CAShapeLayer()

    override init(frame: CGRect) {
        super.init(frame: frame)
        setup()
    }

    required init?(coder aDecoder: NSCoder) {
        super.init(coder: aDecoder)
        setup()
    }

    private func setup() {
        borderLayer.lineWidth = 3
        borderLayer.fillColor = UIColor.clear.cgColor
        borderLayer.strokeColor = UIColor.blue.cgColor
        borderLayer.lineJoin = kCALineJoinRound
        layer.addSublayer(borderLayer)

        [leftCornerHandle, rightCornerHandle].forEach { (handle) in
            handle.backgroundColor = .blue
            handle.frame.size = CGSize(width: 12, height: 12)
            handle.layer.cornerRadius = 6
            handle.layer.borderWidth = 2
            handle.layer.borderColor = UIColor.white.cgColor
            handle.layer.shadowColor = UIColor.black.cgColor
            handle.layer.shadowOpacity = 0.5
            handle.layer.shadowRadius = 1
            handle.layer.shadowOffset = CGSize(width: 0, height: 1)
        }
        addSubview(leftCornerHandle)
        addSubview(rightCornerHandle)
    }

    override func layoutSubviews() {
        super.layoutSubviews()
        let path = UIBezierPath(rect: bounds)
        path.lineJoinStyle = .round

        borderLayer.frame = bounds
        borderLayer.path = path.cgPath

        leftCornerHandle.center = CGPoint(x: 0, y: 0)
        rightCornerHandle.center = CGPoint(x: bounds.maxX, y: bounds.maxY)
    }
}

class MergedCellStore {
    var mergedCells = [IndexPath: CellRange]()

    subscript(_ indexPath: IndexPath) -> CellRange? {
        get {
            return mergedCells[indexPath]
        }
        set {
            return mergedCells[indexPath] = newValue
        }
    }

    func intersection(cellRange: CellRange) -> CellRange? {
        for indexPath in cellRange {
            if let existingMergedCell = self[indexPath] {
                if let intersection = cellRange.intersection(existingMergedCell) {
                    return intersection
                }
            }
        }
        return nil
    }
}

class TextCell: Cell {
    let label = UILabel()
    var font = UIFont.systemFont(ofSize: 12)
    var textAlignment: NSTextAlignment = .left
    var text: String = "" {
        didSet {
            label.text = text
        }
    }
    var attributedText: NSAttributedString = NSAttributedString() {
        didSet {
            label.attributedText = attributedText
        }
    }

    override init(frame: CGRect) {
        super.init(frame: frame)
        setup()
    }

    required init?(coder aDecoder: NSCoder) {
        super.init(coder: aDecoder)
        setup()
    }

    private func setup() {
        label.frame = bounds
        label.autoresizingMask = [.flexibleWidth, .flexibleHeight]
        label.font = font
        label.textAlignment = textAlignment
        contentView.addSubview(label)
    }

    override func layoutSubviews() {
        super.layoutSubviews()
        label.frame = bounds.insetBy(dx: 2, dy: 2)
    }
}

class DebugCell: Cell {
    let label = UILabel()
    var indexPath: IndexPath! {
        didSet {
            label.text = "R\(indexPath.row)C\(indexPath.column)"
        }
    }

    override init(frame: CGRect) {
        super.init(frame: frame)
        setup()
    }

    required init?(coder aDecoder: NSCoder) {
        super.init(coder: aDecoder)
        setup()
    }

    private func setup() {
        label.frame = bounds
        label.autoresizingMask = [.flexibleWidth, .flexibleHeight]
        label.font = UIFont.systemFont(ofSize: 10)
        label.textAlignment = .center
        contentView.addSubview(label)
    }
}

struct SelectionRange: Sequence {
    let from: IndexPath
    let to: IndexPath

    var normalizedSelectionRange: SelectionRange {
        if to.column < from.column && to.row < from.row {
            return SelectionRange(from: to, to: from)
        } else if to.column < from.column {
            return SelectionRange(from: IndexPath(row: from.row, column: to.column), to: IndexPath(row: to.row, column: from.column))
        } else  if to.row < from.row {
            return SelectionRange(from: IndexPath(row: to.row, column: from.column), to: IndexPath(row: from.row, column: to.column))
        } else {
            return SelectionRange(from: from, to: to)
        }
    }

    func expandSelectionRange(mergedCells: [IndexPath: CellRange]) -> SelectionRange {
        var from = (row: self.from.row, column: self.from.column)
        var to = (row: self.to.row, column: self.to.column)
        for indexPath in self {
            if let mergedCell = mergedCells[indexPath] {
                if from.column > mergedCell.from.column {
                    from.column = mergedCell.from.column
                }
                if from.row > mergedCell.from.row {
                    from.row = mergedCell.from.row
                }
                if to.column < mergedCell.to.column {
                    to.column = mergedCell.to.column
                }
                if to.row < mergedCell.to.row {
                    to.row = mergedCell.to.row
                }
            }
        }
        return SelectionRange(from: IndexPath(row: from.row, column: from.column), to: IndexPath(row: to.row, column: to.column))
    }

    public typealias Iterator = SelectionRangeIterator

    public func makeIterator() -> SelectionRangeIterator {
        return SelectionRangeIterator(selectionRange: self)
    }

    struct SelectionRangeIterator: IteratorProtocol {
        public typealias Element = IndexPath

        private let selectionRange: SelectionRange
        private var column: Int
        private var row: Int

        init(selectionRange: SelectionRange) {
            self.selectionRange = selectionRange
            column = selectionRange.from.column
            row = selectionRange.from.row
        }

        public mutating func next() -> IndexPath? {
            if column > selectionRange.to.column {
                column = 0
                row += 1
                if row > selectionRange.to.row {
                    return nil
                }
            }
            let indexPath = IndexPath(row: row, column: column)
            column += 1
            return indexPath
        }
    }
}

public class Spreadsheet: UIView, SpreadsheetViewDelegate, UIGestureRecognizerDelegate, UITextFieldDelegate {
    public var delegate: SpreadsheetDelegate?

    private let spreadsheetDataSource = SpreadsheetDataSource()

    private let spreadsheetView = SpreadsheetView()
    private let selectionView = SelectionView()

    private var selectionRange: SelectionRange?
    private var selectedCellRange: CellRange?

    var isLeftHandleDragging = false
    var isRightHandleDragging = false

    var isCellDragging = false
    var previousLocation: CGPoint = .zero
    var previousIndexPath = IndexPath(row: 0, column: 0)
    var destinationSelectionRange: SelectionRange?
    var snapshotView = UIView()
    var sourceView = UIView()

    let textField = UITextField()
    var editingIndexPath: IndexPath?

    private var isMenuVisible: Bool {
        let menuController = UIMenuController.shared
        return menuController.isMenuVisible
    }

    public override init(frame: CGRect) {
        super.init(frame: frame)
        setup()
    }
    
    public required init?(coder aDecoder: NSCoder) {
        super.init(coder: aDecoder)
        setup()
    }

    private func setup() {
        spreadsheetView.frame = bounds
        spreadsheetView.autoresizingMask = [.flexibleWidth, .flexibleHeight]
        spreadsheetView.dataSource = spreadsheetDataSource
        spreadsheetView.delegate = self
        spreadsheetView.register(TextCell.self, forCellWithReuseIdentifier: String(describing: TextCell.self))
        addSubview(spreadsheetView)

        selectionView.isHidden = true
        spreadsheetView.addSubview(selectionView)

        sourceView.backgroundColor = UIColor(white: 0.85, alpha: 1)
        textField.delegate = self
        textField.font = UIFont.systemFont(ofSize: 12)

        spreadsheetView.addGestureRecognizer(UIPanGestureRecognizer(target: self, action: #selector(spreadsheetDragged(_:))))

        let cellLongPressGestureRecognizer = UILongPressGestureRecognizer(target: self, action: #selector(cellLongPressed(_:)))
        cellLongPressGestureRecognizer.minimumPressDuration = 0.25
        spreadsheetView.addGestureRecognizer(cellLongPressGestureRecognizer)

        let cellDoubleTapGestureRecognizer = UITapGestureRecognizer(target: self, action: #selector(cellDoubleTapped(_:)))
        cellDoubleTapGestureRecognizer.numberOfTapsRequired = 2
        spreadsheetView.addGestureRecognizer(cellDoubleTapGestureRecognizer)
    }

    public func textFieldShouldReturn(_ textField: UITextField) -> Bool {
        textField.resignFirstResponder()
        return true
    }

    public func textFieldDidEndEditing(_ textField: UITextField) {
        if let editingIndexPath = editingIndexPath {
            spreadsheetDataSource.data[editingIndexPath] = textField.text ?? ""
            spreadsheetView.reloadData()
        }
        textField.removeFromSuperview()
    }

    func cellDoubleTapped(_ gestureRecognizer: UITapGestureRecognizer) {
        let location = gestureRecognizer.location(in: spreadsheetView)
        if let indexPath = spreadsheetView.indexPathForItem(at: location) {
            selectionRange = SelectionRange(from: indexPath, to: indexPath)
            updateSelectionView(selectionRange: selectionRange!)
            updateSelectedCellRange(selectionRange: selectionRange!)

            if delegate?.spreadsheet(self, textShouldBeginEditingAt: indexPath) ?? true {
                if let cell = spreadsheetView.cellForItem(at: indexPath) as? TextCell {
                    cell.text = ""
                }

                textField.frame = spreadsheetView.rectForItem(at: indexPath).insetBy(dx: 2, dy: 2)
                textField.text = spreadsheetDataSource.data[indexPath]
                spreadsheetView.addSubview(textField)
                textField.becomeFirstResponder()
                delegate?.spreadsheet(self, textDidBeginEditingAt: indexPath)

                editingIndexPath = indexPath
            }
        }
    }

    func cellLongPressed(_ gestureRecognizer: UILongPressGestureRecognizer) {
        let location = gestureRecognizer.location(in: spreadsheetView)

        switch gestureRecognizer.state {
        case .began:
            hideMenu()
            spreadsheetView.isScrollEnabled = false

            if let indexPath = spreadsheetView.indexPathForItem(at: location) {
                if let selectedCellRange = selectedCellRange, selectedCellRange.contains(indexPath) {
                    print(selectedCellRange)
                    print(selectionRange)
                    destinationSelectionRange = selectionRange

                    selectionView.leftCornerHandle.isHidden = true
                    selectionView.rightCornerHandle.isHidden = true
                    selectionView.isHidden = true

                    let frame = selectionView.frame.insetBy(dx: 4, dy: 4)
                    sourceView.frame = frame
                    spreadsheetView.insertSubview(sourceView, belowSubview: selectionView)

                    snapshotView = spreadsheetView.resizableSnapshotView(from: frame, afterScreenUpdates: false, withCapInsets: .zero)!
                    snapshotView.frame = frame
                    spreadsheetView.addSubview(snapshotView)

                    previousLocation = location
                    previousIndexPath = indexPath
                    isCellDragging = true

                    UIView.animate(withDuration: CATransaction.animationDuration(), delay: 0, options: .curveEaseOut, animations: {
                        self.snapshotView.frame.origin.y -= 8
                        self.snapshotView.alpha = 0.5
                        self.snapshotView.layer.shadowColor = UIColor.black.cgColor
                        self.snapshotView.layer.shadowOffset = .zero
                        self.snapshotView.layer.shadowOpacity = 0.5
                        self.snapshotView.layer.shadowRadius = 4
                    }, completion: nil)
                } else {
                    selectionRange = SelectionRange(from: indexPath, to: indexPath)
                    updateSelectionView(selectionRange: selectionRange!)
                    updateSelectedCellRange(selectionRange: selectionRange!)
                    selectionView.isHidden = false
                }
            }
        case .changed:
            if isCellDragging {
                if let indexPath = spreadsheetView.indexPathForItem(at: location), let selectedCellRange = selectedCellRange {
                    let diffColumn = previousIndexPath.column - indexPath.column
                    let diffRow = previousIndexPath.row - indexPath.row

                    var fromRow = selectedCellRange.from.row - diffRow
                    var fromColumn = selectedCellRange.from.column - diffColumn
                    var toRow = selectedCellRange.to.row - diffRow
                    var toColumn = selectedCellRange.to.column - diffColumn
                    if fromRow < 0 || toRow > spreadsheetView.numberOfRows || fromColumn < 0 || toColumn > spreadsheetView.numberOfColumns {
                        fromRow = destinationSelectionRange!.from.row
                        fromColumn = destinationSelectionRange!.from.column
                        toRow = destinationSelectionRange!.to.row
                        toColumn = destinationSelectionRange!.to.column
                    }

                    destinationSelectionRange = SelectionRange(from: IndexPath(row: fromRow, column: fromColumn), to: IndexPath(row: toRow, column: toColumn))
                    updateSelectionView(selectionRange: destinationSelectionRange!)
                    selectionView.isHidden = false
                }
                let diffX = previousLocation.x - location.x
                let diffY = previousLocation.y - location.y
                let center = snapshotView.center
                snapshotView.center = CGPoint(x: center.x - diffX, y: center.y - diffY)
                previousLocation = location
            } else if let indexPath = spreadsheetView.indexPathForItem(at: location), let selectionRange = selectionRange {
                self.selectionRange = SelectionRange(from: selectionRange.from, to: indexPath)
                self.selectionRange = self.selectionRange?.normalizedSelectionRange
                self.selectionRange = self.selectionRange?.expandSelectionRange(mergedCells: spreadsheetDataSource.mergedCellStore.mergedCells)
                updateSelectionView(selectionRange: self.selectionRange!)
            }
        case .ended:
            spreadsheetView.isScrollEnabled = true

            if isCellDragging {
                var sourceData = [Int: String]()
                var index = 0
                if let selectionRange = selectionRange {
                    for column in selectionRange.from.column...selectionRange.to.column {
                        for row in selectionRange.from.row...selectionRange.to.row {
                            let indexPath = IndexPath(row: row, column: column)
                            sourceData[index] = spreadsheetDataSource.data[indexPath]
                            spreadsheetDataSource.data[indexPath] = nil
                            index += 1
                        }
                    }
                }
                index = 0
                if let selectionRange = destinationSelectionRange {
                    for column in selectionRange.from.column...selectionRange.to.column {
                        for row in selectionRange.from.row...selectionRange.to.row {
                            let indexPath = IndexPath(row: row, column: column)
                            spreadsheetDataSource.data[indexPath] = sourceData[index]
                            index += 1
                        }
                    }
                }
                spreadsheetView.reloadData()

                UIView.animate(withDuration: CATransaction.animationDuration(), delay: 0, options: .curveEaseOut, animations: {
                    self.snapshotView.frame = self.selectionView.frame.insetBy(dx: 4, dy: 4)
                    self.snapshotView.alpha = 1
                    self.snapshotView.layer.shadowOpacity = 0
                    self.snapshotView.layer.shadowRadius = 0
                    self.sourceView.alpha = 0
                }, completion: { (finished) in
                    self.snapshotView.removeFromSuperview()
                    self.sourceView.removeFromSuperview()
                    self.sourceView.alpha = 1
                })

                selectionRange = destinationSelectionRange
                updateSelectedCellRange(selectionRange: selectionRange!)

                isCellDragging = false
                selectionView.leftCornerHandle.isHidden = false
                selectionView.rightCornerHandle.isHidden = false
                selectionView.isHidden = false
            } else {
                if let selectionRange = selectionRange {
                    updateSelectedCellRange(selectionRange: selectionRange)
                    showMenu()
                }
            }
        default:
            spreadsheetView.isScrollEnabled = true
            isCellDragging = false
            selectionView.leftCornerHandle.isHidden = false
            selectionView.rightCornerHandle.isHidden = false
            snapshotView.removeFromSuperview()
        }
    }

    func spreadsheetDragged(_ gestureRecognizer: UIPanGestureRecognizer) {
        let location = gestureRecognizer.location(in: spreadsheetView)
        switch gestureRecognizer.state {
        case .began:
            hideMenu()
            if selectionView.convert(CGRect(x: -20, y: -20, width: 40, height: 40), to: spreadsheetView).contains(location) {
                spreadsheetView.isScrollEnabled = false
                isLeftHandleDragging = true
            }
            if selectionView.convert(CGRect(x: selectionView.bounds.maxX - 20, y: selectionView.bounds.maxY - 20, width: 40, height: 40), to: spreadsheetView).contains(location) {
                spreadsheetView.isScrollEnabled = false
                isRightHandleDragging = true
            }
        case .changed:
            if isLeftHandleDragging {
                if let indexPath = spreadsheetView.indexPathForItem(at: location), let selectionRange = selectionRange {
                    self.selectionRange = SelectionRange(from: indexPath, to: selectionRange.to)
                    self.selectionRange = self.selectionRange?.normalizedSelectionRange
                    self.selectionRange = self.selectionRange?.expandSelectionRange(mergedCells: spreadsheetDataSource.mergedCellStore.mergedCells)
                    updateSelectionView(selectionRange: self.selectionRange!)
                }
            }
            if isRightHandleDragging {
                if let indexPath = spreadsheetView.indexPathForItem(at: location), let selectionRange = selectionRange {
                    self.selectionRange = SelectionRange(from: selectionRange.from, to: indexPath)
                    self.selectionRange = self.selectionRange?.normalizedSelectionRange
                    self.selectionRange = self.selectionRange?.expandSelectionRange(mergedCells: spreadsheetDataSource.mergedCellStore.mergedCells)
                    updateSelectionView(selectionRange: self.selectionRange!)
                }
            }
        case .ended:
            if isLeftHandleDragging || isRightHandleDragging {
                if let selectionRange = selectionRange {
                    showMenu()
                    updateSelectedCellRange(selectionRange: selectionRange)
                }
            }
            spreadsheetView.isScrollEnabled = true
            isLeftHandleDragging = false
            isRightHandleDragging = false
        default:
            spreadsheetView.isScrollEnabled = true
            isLeftHandleDragging = false
            isRightHandleDragging = false
        }
    }

    func cellAction(_ sender: Any?) {
        guard let delegate = delegate else {
            return
        }
        if let cellRange = selectedCellRange {
            let intersection = spreadsheetDataSource.mergedCellStore.intersection(cellRange: cellRange)
            delegate.spreadsheet(self, performCellAction: cellRange, intersection: intersection)
        }
    }

    public func mergeCells(cellRange: CellRange) {
        spreadsheetDataSource.mergeCells(cellRange: cellRange)
        spreadsheetView.reloadData()
    }

    public func unmergeCell(cellRange: CellRange) {
        spreadsheetDataSource.unmergeCell(cellRange: cellRange)
        spreadsheetView.reloadData()
    }

    public override func paste(_ sender: Any?) {

    }

    public override func canPerformAction(_ action: Selector, withSender sender: Any?) -> Bool {
        return action == #selector(paste(_:)) || action == #selector(cellAction(_:))
    }

    public override var canBecomeFirstResponder: Bool {
        return true
    }

    private func updateSelectionView(selectionRange: SelectionRange) {
        let fromRect = spreadsheetView.rectForItem(at: selectionRange.from)
        let toRect = spreadsheetView.rectForItem(at: selectionRange.to)
        selectionView.frame.origin = CGPoint(x: fromRect.origin.x, y: fromRect.origin.y)
        selectionView.frame.size = CGSize(width: toRect.maxX - fromRect.minX, height: toRect.maxY - fromRect.minY )
        selectionView.frame = selectionView.frame.insetBy(dx: -4, dy: -4)
    }

    private func updateSelectedCellRange(selectionRange: SelectionRange) {
        selectedCellRange = CellRange(from: selectionRange.from, to: selectionRange.to)
    }

    private func showMenu() {
        becomeFirstResponder()
        let menuController = UIMenuController.shared
        menuController.menuItems = [UIMenuItem(title: "Cell Actions...", action: #selector(cellAction(_:)))]
        menuController.setTargetRect(selectionView.frame, in: spreadsheetView)
        menuController.setMenuVisible(true, animated: true)
    }

    private func hideMenu() {
        let menuController = UIMenuController.shared
        menuController.setMenuVisible(false, animated: true)
        resignFirstResponder()
    }

    public func spreadsheetView(_ spreadsheetView: SpreadsheetView, didSelectItemAt indexPath: IndexPath) {
        if let range = selectionRange {
            let cellRange = CellRange(from: range.from, to: range.to)
            if cellRange.contains(indexPath) {
                if isMenuVisible {
                    hideMenu()
                } else {
                    showMenu()
                }
            } else {
                hideMenu()
                selectionRange = SelectionRange(from: indexPath, to: indexPath)
                updateSelectionView(selectionRange: selectionRange!)
                updateSelectedCellRange(selectionRange: selectionRange!)
            }
        } else {
            hideMenu()
            selectionRange = SelectionRange(from: indexPath, to: indexPath)
            updateSelectionView(selectionRange: selectionRange!)
            updateSelectedCellRange(selectionRange: selectionRange!)
        }
        selectionView.isHidden = false
    }
}

extension CellRange {
    func isEqual(cellRange: CellRange) -> Bool {
        return self == cellRange && self.to.column == cellRange.to.column && self.to.row == cellRange.to.row
    }
}

extension CellRange: Sequence {
    public typealias Iterator = CellRangeIterator

    public func makeIterator() -> CellRangeIterator {
        return CellRangeIterator(cellRange: self)
    }
}

public struct CellRangeIterator: IteratorProtocol {
    public typealias Element = IndexPath

    private let cellRange: CellRange
    private var column: Int
    private var row: Int

    init(cellRange: CellRange) {
        self.cellRange = cellRange
        column = cellRange.from.column
        row = cellRange.from.row
    }

    public mutating func next() -> IndexPath? {
        if column > cellRange.to.column {
            column = 0
            row += 1
            if row > cellRange.to.row {
                return nil
            }
        }
        let indexPath = IndexPath(row: row, column: column)
        column += 1
        return indexPath
    }
}

class ViewController: UIViewController, SpreadsheetDelegate {
    let spreadsheet = Spreadsheet()

    override func viewDidLoad() {
        super.viewDidLoad()

        spreadsheet.frame = view.bounds
        spreadsheet.autoresizingMask = [.flexibleWidth, .flexibleHeight]
        spreadsheet.delegate = self
        view.addSubview(spreadsheet)
    }

    func spreadsheet(_ spreadsheet: Spreadsheet, performCellAction cellRange: CellRange, intersection: CellRange?) {
        let controller = UIAlertController(title: "\(cellRange)", message: nil, preferredStyle: .actionSheet)
        if let intersection = intersection {
            if cellRange.isEqual(cellRange: intersection) {
                controller.addAction(UIAlertAction(title: NSLocalizedString("Unmerge Cells", comment: ""), style: .default) { (action) in
                    spreadsheet.unmergeCell(cellRange: cellRange)
                })
            } else {
                controller.addAction(UIAlertAction(title: NSLocalizedString("Merge All", comment: ""), style: .default) { (action) in
                    spreadsheet.mergeCells(cellRange: cellRange)
                })
                controller.addAction(UIAlertAction(title: NSLocalizedString("Unmerge All", comment: ""), style: .default) { (action) in
                    spreadsheet.unmergeCell(cellRange: cellRange)
                })
            }
        } else if cellRange.columnCount == 1 && cellRange.rowCount == 1 {

        } else {
            controller.addAction(UIAlertAction(title: NSLocalizedString("Merge Cells", comment: ""), style: .default) { (action) in
                spreadsheet.mergeCells(cellRange: cellRange)
            })
        }
        controller.addAction(UIAlertAction(title: NSLocalizedString("Cancel", comment: ""), style: .cancel))
        present(controller, animated: true, completion: nil)
    }
}
