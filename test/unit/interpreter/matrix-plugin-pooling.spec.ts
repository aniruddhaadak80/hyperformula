import {CellError, DetailedCellError, ErrorType, HyperFormula} from '../../../src'
import {SimpleCellAddress, simpleCellAddress} from '../../../src/Cell'
import {Config} from '../../../src/Config'
import {ErrorMessage} from '../../../src/error-message'

const colNumber = (input: string): number =>
  input.split('').reduce((currentColumn, nextLetter) => currentColumn * 26 + (nextLetter.toUpperCase().charCodeAt(0) - 64), 0) - 1

const adr = (stringAddress: string, sheet: number = 0): SimpleCellAddress => {
  const result = /^([A-Za-z]+)([0-9]+)$/.exec(stringAddress)!
  return simpleCellAddress(sheet, colNumber(result[1]), Number(result[2]) - 1)
}

const detailedError = (errorType: ErrorType, message?: string): DetailedCellError =>
  new DetailedCellError(new CellError(errorType, message), new Config().translationPackage.getErrorTranslation(errorType))

const poolDimensionsError = () => detailedError(ErrorType.VALUE, ErrorMessage.PoolDimensions)

const POOL_FUNCTION_NAMES = ['MAXPOOL', 'MEDIANPOOL']

const squareRange3x3 = [
  [3, 1, 2],
  [9, 7, 8],
  [5, 4, 6],
]

const squareRange4x4 = [
  [1, 2, 10, 20],
  [3, 4, 30, 40],
  [5, 6, 7, 8],
  [7, 8, 9, 10],
]

const buildEngineWithSheet = (sheetContent: number[][]): { engine: HyperFormula, sheetId: number } => {
  const engine = HyperFormula.buildEmpty()
  const sheetId = engine.getSheetId(engine.addSheet('Sheet1'))!
  engine.setSheetContent(sheetId, sheetContent)
  return {engine, sheetId}
}

// A plain loop instead of `describe.each`, which is not available in the Jasmine-based browser test suite.
POOL_FUNCTION_NAMES.forEach((functionName: string) => describe(`Function ${functionName} dimension validation`, () => {
  it('returns an error when the range dimensions are not a whole multiple of the window size', () => {
    const {engine, sheetId} = buildEngineWithSheet(squareRange3x3)

    expect(engine.calculateFormula(`=${functionName}(A1:C3, 2)`, sheetId)).toEqualError(poolDimensionsError())
  })

  it('returns an error in a cell when the range dimensions are not a whole multiple of the window size', () => {
    const {engine, sheetId} = buildEngineWithSheet(squareRange3x3)

    engine.setCellContents(adr('A5', sheetId), [[`=${functionName}(A1:C3, 2)`]])

    expect(engine.getCellValue(adr('A5', sheetId))).toEqualError(poolDimensionsError())
  })

  it('returns an error when only one of the range dimensions is a whole multiple of the window size', () => {
    const {engine, sheetId} = buildEngineWithSheet([
      [1, 2, 3, 4],
      [5, 6, 7, 8],
      [9, 10, 11, 12],
    ])

    expect(engine.calculateFormula(`=${functionName}(A1:D3, 2)`, sheetId)).toEqualError(poolDimensionsError())
  })

  it('returns an error when the window is larger than the range', () => {
    const {engine, sheetId} = buildEngineWithSheet(squareRange3x3)

    expect(engine.calculateFormula(`=${functionName}(A1:C3, 4)`, sheetId)).toEqualError(poolDimensionsError())
  })

  it('returns an error when the stride makes the last window reach outside of the range', () => {
    const {engine, sheetId} = buildEngineWithSheet(squareRange3x3)

    expect(engine.calculateFormula(`=${functionName}(A1:C3, 2, 3)`, sheetId)).toEqualError(poolDimensionsError())
  })

  it('returns an error when the window size is not a positive integer', () => {
    const {engine, sheetId} = buildEngineWithSheet(squareRange4x4)

    expect(engine.calculateFormula(`=${functionName}(A1:D4, 0)`, sheetId)).toEqualError(detailedError(ErrorType.NUM, ErrorMessage.ValueSmall))
    expect(engine.calculateFormula(`=${functionName}(A1:D4, -2)`, sheetId)).toEqualError(detailedError(ErrorType.NUM, ErrorMessage.ValueSmall))
    expect(engine.calculateFormula(`=${functionName}(A1:D4, 2.5)`, sheetId)).toEqualError(detailedError(ErrorType.NUM, ErrorMessage.IntegerExpected))
  })

  it('returns an error when the stride is not a positive integer', () => {
    const {engine, sheetId} = buildEngineWithSheet(squareRange4x4)

    expect(engine.calculateFormula(`=${functionName}(A1:D4, 2, 0)`, sheetId)).toEqualError(detailedError(ErrorType.NUM, ErrorMessage.ValueSmall))
    expect(engine.calculateFormula(`=${functionName}(A1:D4, 2, -1)`, sheetId)).toEqualError(detailedError(ErrorType.NUM, ErrorMessage.ValueSmall))
    expect(engine.calculateFormula(`=${functionName}(A1:D4, 2, 1.5)`, sheetId)).toEqualError(detailedError(ErrorType.NUM, ErrorMessage.IntegerExpected))
  })
}))

describe('Function MAXPOOL', () => {
  it('pools a range whose dimensions are a whole multiple of the window size into maximums', () => {
    const {engine, sheetId} = buildEngineWithSheet(squareRange4x4)

    expect(engine.calculateFormula('=MAXPOOL(A1:D4, 2)', sheetId)).toEqual([
      [4, 40],
      [8, 10],
    ])
  })

  it('pools a range whose dimensions are a whole multiple of a custom stride', () => {
    const {engine, sheetId} = buildEngineWithSheet(squareRange3x3)

    expect(engine.calculateFormula('=MAXPOOL(A1:C3, 2, 1)', sheetId)).toEqual([
      [9, 8],
      [9, 8],
    ])
  })

  it('pools a range with a window covering the whole range', () => {
    const {engine, sheetId} = buildEngineWithSheet(squareRange3x3)

    expect(engine.calculateFormula('=MAXPOOL(A1:C3, 3)', sheetId)).toEqual(9)
  })

  it('spills the pooled maximums into cells', () => {
    const {engine, sheetId} = buildEngineWithSheet(squareRange4x4)

    engine.setCellContents(adr('A6', sheetId), [['=MAXPOOL(A1:D4, 2)']])

    expect(engine.getCellValue(adr('A6', sheetId))).toEqual(4)
    expect(engine.getCellValue(adr('B6', sheetId))).toEqual(40)
    expect(engine.getCellValue(adr('A7', sheetId))).toEqual(8)
    expect(engine.getCellValue(adr('B7', sheetId))).toEqual(10)
  })
})

describe('Function MEDIANPOOL', () => {
  it('pools a range whose dimensions are a whole multiple of the window size into medians', () => {
    const {engine, sheetId} = buildEngineWithSheet(squareRange4x4)

    const result = engine.calculateFormula('=MEDIANPOOL(A1:D4, 2)', sheetId) as number[][]

    expect(result[0][0]).toBeCloseTo(2.5)
    expect(result[0][1]).toBeCloseTo(25)
    expect(result[1][0]).toBeCloseTo(6.5)
    expect(result[1][1]).toBeCloseTo(8.5)
  })

  it('spills the pooled medians into cells', () => {
    const {engine, sheetId} = buildEngineWithSheet(squareRange4x4)

    engine.setCellContents(adr('A6', sheetId), [['=MEDIANPOOL(A1:D4, 2)']])

    expect(engine.getCellValue(adr('A6', sheetId))).toBeCloseTo(2.5)
    expect(engine.getCellValue(adr('B6', sheetId))).toBeCloseTo(25)
    expect(engine.getCellValue(adr('A7', sheetId))).toBeCloseTo(6.5)
    expect(engine.getCellValue(adr('B7', sheetId))).toBeCloseTo(8.5)
  })
})
