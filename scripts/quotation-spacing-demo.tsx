import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { ResponsiveGrid } from '../CMModelDrivenGrid/ResponsiveGrid';
import '../CMModelDrivenGrid/css/CMModelDrivenGrid.css';

initializeIcons(undefined,{disableWarnings:true});
const optionNames=['Hardware Colour','Barrel Type','Mechanism','Fixing Option','Bracket Cover','Chain','Control Side','Chain Connectors','Bottom Bar','Chain/Cord Tidies','Roll Direction','Production Class','Branded Labels'];
const choices=[{id:'grey',name:'H-AGREY | Anthracite Grey Hardware Colour'},{id:'black',name:'H-BLACK | Black Hardware Colour'},{id:'white',name:'H-WHITE | White Hardware Colour'}];
const columns=[
    {name:'order',displayName:'Order',dataType:'Whole.None',visualSizeFactor:60},
    {name:'option',displayName:'Option',dataType:'SingleLine.Text',visualSizeFactor:180,isReadOnly:true},
    {name:'component',displayName:'Component',dataType:'Lookup.Simple',visualSizeFactor:320,targets:['cm_component']},
    {name:'calculator',displayName:'Calculator',dataType:'SingleLine.Text',visualSizeFactor:110},
    {name:'cost',displayName:'Total Cost',dataType:'Currency',visualSizeFactor:100},
].map((column,order)=>({...column,order,alias:column.name})) as ComponentFramework.PropertyHelper.DataSetApi.Column[];
const resources={getString:(key:string)=>({'Label_Grid_Footer_RecordCount':'{0} records ({1} Selected)','Label_NoRecords':'No records found','Label_Grid_Footer':'Page {0}'}[key]??key),getResource:()=>undefined} as ComponentFramework.Resources;
type DemoRow={id:string;order:number;option:string;component:{id:string;name:string;entityType:string}|null;calculator:string;cost:number};
const sampleRows=():DemoRow[]=>optionNames.map((option,order)=>({id:`row-${order}`,order,option,component:order===1?{id:'tube',name:'R8024 | 40mm Roller Tube',entityType:'cm_component'}:null,calculator:order===0?'':order===1||order===8?'Width':order===5?'Drop':'Quantity',cost:order===2?2.59:0}));
function App(): React.ReactElement {
    const [rows,setRows]=React.useState<DemoRow[]>(sampleRows);
    const [maximum,setMaximum]=React.useState(900),[fit,setFit]=React.useState(true),[selected,setSelected]=React.useState<string[]>([]),[writes,setWrites]=React.useState(0),[status,setStatus]=React.useState('Sample data only');
    const [sorts,setSorts]=React.useState<ComponentFramework.PropertyHelper.DataSetApi.SortStatus[]>([]);
    const records=Object.fromEntries(rows.map(row=>[row.id,{
        getRecordId:()=>row.id,getValue:(key:string)=>(row as unknown as Record<string,unknown>)[key],
        getFormattedValue:(key:string)=>key==='component'?row.component?.name??'':key==='cost'?`£${row.cost.toFixed(2)}`:String((row as unknown as Record<string,unknown>)[key]??''),
        getNamedReference:()=>({id:{guid:row.id},etn:'cm_quotationlineoption',name:row.option}),
    } as ComponentFramework.PropertyHelper.DataSetApi.EntityRecord]));
    const ordered=[...rows];if(sorts.length){const sort=sorts[0];ordered.sort((a,b)=>String((a as unknown as Record<string,unknown>)[sort.name]??'').localeCompare(String((b as unknown as Record<string,unknown>)[sort.name]??''))*(sort.sortDirection===1?-1:1));}
    return <>
        <header><strong>Quotation grid spacing · local harness</strong><span>{status} · Saves: {writes}</span></header>
        <div className="tools"><button onClick={()=>setMaximum(900)}>Tall area</button><button onClick={()=>setMaximum(360)}>Short area</button><label><input type="checkbox" checked={fit} onChange={event=>setFit(event.target.checked)}/> Fit to rows</label><button onClick={()=>setRows(current=>[...current,...Array.from({length:40},(_,index)=>({id:`extra-${current.length+index}`,order:current.length+index,option:`Extra option ${current.length+index}`,component:null,calculator:'Quantity',cost:0}))])}>Add sample rows</button><button onClick={()=>setRows([])}>Clear sample rows</button><button onClick={()=>setRows(sampleRows())}>Restore sample rows</button></div>
        <main><div className="commands"><span>{selected.length} Selected</span><button onClick={()=>setStatus('Refresh selected')}>Refresh</button></div>
        <ResponsiveGrid width={undefined} height={maximum} fitToRows={fit} columns={columns} records={records} sortedRecordIds={ordered.map(row=>row.id)}
            hasNextPage={false} hasPreviousPage={false} totalResultCount={rows.length} currentPage={1} sorting={sorts} filtering={undefined as unknown as ComponentFramework.PropertyHelper.DataSetApi.FilterExpression} resources={resources} itemsLoading={false}
            highlightValue={null} highlightColor={null} enableLookupLinks={true} enableInlineEdit={true} infiniteScroll={true}
            childRows={{loadPage:async()=>({records:[],hasNextPage:false})}}
            lookupTargetEntityMap={{component:'cm_component'}} onSearchLookup={async(_entity,term)=>choices.filter(choice=>choice.name.toLowerCase().includes(term.toLowerCase()))}
            onBrowseLookup={async()=>{setStatus('Full lookup requested');return undefined;}}
            onUpdateLookupCell={async(id,key,value)=>{setWrites(count=>count+1);setRows(current=>current.map(row=>row.id===id?{...row,[key]:value?{...choices.find(choice=>choice.id===value)!,entityType:'cm_component'}:null}:row));}}
            setSelectedRecords={setSelected} onNavigate={()=>undefined} onOpenLookup={()=>undefined}
            onSort={(name,direction)=>setSorts(direction==='none'?[]:[{name,sortDirection:direction==='desc'?1:0}])}
            onFilter={()=>setStatus('Filter requested')} onClearAllFilters={()=>undefined} loadFirstPage={()=>undefined} loadNextPage={()=>undefined} loadPreviousPage={()=>undefined}
            onUpdateCell={async(id,key,value)=>{setWrites(count=>count+1);setRows(current=>current.map(row=>row.id===id?{...row,[key]:value}:row));}}/>
        <div className="next-section">Next form section</div></main>
    </>;
}
ReactDOM.render(<App/>,document.getElementById('spacing-root'));
