import { describe, expect, it } from "vitest";
import { findCells } from "../find-engine.js";
import type { WorksheetView } from "@opensheet/core";
const sheet = (cells: Record<string, { value: any; formula?: string }>): WorksheetView => ({ id:"s",name:"s",rowCount:10,columnCount:10,frozenRows:0,frozenColumns:0,cellCount:Object.keys(cells).length,filter:null,getCell:(r,c)=>cells[`${r}:${c}`],*cellEntries(){for(const [key,data] of Object.entries(cells)){const [row,col]=key.split(":").map(Number);yield [row!,col!,data] as any;}},getRowHeight:()=>undefined,getColumnWidth:()=>undefined,forEachCellInRange:()=>{} } as WorksheetView);
const options = { query:"a",matchCase:false,wholeCell:false,searchIn:"values" as const,scope:"all" as const,direction:"forward" as const };
describe("findCells",()=>{
 it("uses sparse deterministic row-major traversal and text matching",()=>expect(findCells(sheet({"2:1":{value:"Alpha"},"0:2":{value:"beta"},"0:1":{value:"A"}}),options)).toEqual([{row:0,col:1},{row:0,col:2},{row:2,col:1}]));
 it("distinguishes computed values from formula sources",()=>{const s=sheet({"0:0":{value:20,formula:"=SUM(A1:A2)"}});expect(findCells(s,{...options,query:"20",wholeCell:true})).toEqual([{row:0,col:0}]);expect(findCells(s,{...options,query:"sum",searchIn:"formulas"})).toEqual([{row:0,col:0}]);});
 it("formats booleans and errors and supports reverse traversal",()=>expect(findCells(sheet({"0:0":{value:true},"1:0":{value:{type:"#REF!"}}}),{...options,query:"",direction:"backward"})).toEqual([{row:1,col:0},{row:0,col:0}]));
});
