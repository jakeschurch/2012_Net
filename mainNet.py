#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""copyright Jake Schurch 2018."""
__author__ = "Jake Schurch"

import xlrd


class Edge:
    def __init__(self, sourceNode, targetNode=None, weight=None):
        self.sourceNode = sourceNode.encode('utf-8')
        self.targetNode = targetNode.encode('utf-8')
        self.weight = weight

    def HasNone(self):
        if (self.sourceNode == 'none' or self.targetNode == 'none'
                or self.weight == 'none'):
            return True
        else:
            return False


def GetVerticesMap(EdgeList):
    nodes = []
    for edge in EdgeList:
        nodes.append(edge.sourceNode.decode('utf-8'))
        nodes.append(edge.targetNode.decode('utf-8'))
    nodes = list(set(nodes))
    nodes.sort()
    nodeVertices = {k: v for k, v in enumerate(nodes) if 'none' not in v}
    return nodeVertices


def MapVerticesToEdges(EdgeList, VerticesDict):
    verticesMapping = {v: k for k, v in VerticesDict.items()}
    for edge in EdgeList:
        edge.sourceNode = verticesMapping[edge.sourceNode.decode('utf-8')]
        edge.targetNode = verticesMapping[edge.targetNode.decode('utf-8')]
    return EdgeList


def WriteToOutput(NodeVertices, EdgeList, StateName):
    outputName = '{0} {1} Map.net'.format(nodeSheets[0].name[0:4], StateName)
    with open(outputName, 'w') as file:
        file.write('*Vertices {0}\n'.format(len(NodeVertices)))
        file.writelines('{0} "{1}"\n'.format(k, v)
                        for k, v in NodeVertices.items())

        file.write('*Arcs :1 Appointment\n')
        file.writelines('{0} {1} {2}\n'.format(edge.sourceNode,
                                               edge.targetNode,
                                               edge.weight)
                        for edge in EdgeList)
        file.close()


def main():
    workbook = input('Please enter a filename:\t')
    try:
        wb = xlrd.open_workbook(workbook)
    except FileNotFoundError:
        print('Incorrect file name, please try again.')
        main()

    for sheetName in wb.sheet_names():
        if 'Party' in sheetName:
            if 'Weight' not in sheetName:
                nodeSheets.append(wb.sheet_by_name(sheetName))
            else:
                weightSheets.append(wb.sheet_by_name(sheetName))

    iRow = 1  # skip over header row
    while iRow < 51:
        sheetMap = {k: v for k, v in zip(nodeSheets, weightSheets)}

        edges = [Edge(k.cell_value(iRow, col),
                      k.cell_value(0, col),
                      v.cell_value(iRow, col)
                      )
                 for k, v in sheetMap.items()
                 for col in range(k.ncols) if col is not 0]

        nodeVertices = GetVerticesMap(edges)

        edges = [edge for edge in edges if edge.HasNone() is False]
        edges = MapVerticesToEdges(edges, nodeVertices)

        WriteToOutput(nodeVertices, edges, nodeSheets[0].cell_value(iRow, 0))
        iRow += 1


if __name__ == '__main__':
    nodeSheets = []
    weightSheets = []
    main()
