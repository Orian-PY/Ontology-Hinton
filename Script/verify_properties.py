from pathlib import Path
import sys,json
sys.path.insert(0,'/private/tmp/hinton-owl-check-deps')
from rdflib import Graph,Namespace,RDF,RDFS,OWL,URIRef,Literal
from rdflib.collection import Collection
from owlrl import DeductiveClosure,OWLRL_Semantics

WORK=Path('/private/tmp/hinton-place-properties')
N=Namespace('https://example.org/hinton/newstudent#')
X=Namespace('http://www.w3.org/2001/XMLSchema#')
ERR=Namespace('http://www.daml.org/2002/03/agents/agent-ont#')
audit=json.loads((WORK/'audit.json').read_text())
g=Graph().parse(WORK/'campus_academic_ontology.owl',format='xml')
old=Graph().parse(WORK/'original.owl',format='xml')
checks=[]
def ck(cond,name):
    assert cond,name
    checks.append(name)
def expression(node,predicate):
    values=list(g.objects(node,predicate));ck(len(values)==1,'One '+str(predicate).split('#')[-1]+' for '+str(node).split('#')[-1])
    value=values[0]
    if isinstance(value,URIRef):return {value}
    return set(Collection(g,next(g.objects(value,OWL.unionOf))))
def target(names):
    if isinstance(names,str):names=[names]
    return {N[x] for x in names}

counts={name:len({s for s in g.subjects(RDF.type,term) if isinstance(s,URIRef)}) for name,term in [('classes',OWL.Class),('object_properties',OWL.ObjectProperty),('data_properties',OWL.DatatypeProperty)]}
ck(counts=={'classes':62,'object_properties':30,'data_properties':106},'Named entity counts')
for typ in [OWL.Class,OWL.ObjectProperty,OWL.DatatypeProperty]:
    ck(not {x for x in old.subjects(RDF.type,typ) if isinstance(x,URIRef)}-{x for x in g.subjects(RDF.type,typ) if isinstance(x,URIRef)},'All existing named '+str(typ).split('#')[-1]+' retained')
for name,v in audit['object_new'].items():
    ck((N[name],RDF.type,OWL.ObjectProperty) in g,name+' declared')
    ck(expression(N[name],RDFS.domain)==target(v[0]),name+' domain')
    ck(expression(N[name],RDFS.range)==target(v[1]),name+' range')
for name,v in audit['data_new'].items():
    ck((N[name],RDF.type,OWL.DatatypeProperty) in g,name+' declared')
    ck(expression(N[name],RDFS.domain)==target(v[0]),name+' domain')
    ck(set(g.objects(N[name],RDFS.range))=={X[v[1]]},name+' datatype')
for a,b in audit['aliases'].items():ck((N[a],RDFS.subPropertyOf,N[b]) in g,a+' retained as subproperty of '+b)
ck(expression(N.surfaceType,RDFS.domain)=={N.Room,N.OutdoorSpace},'Surface type domain is one union')
for name in ['hasRoom','hasCampusPlace','locatedInBuilding','locatedOnCampus']:
    ck((N[name],RDF.type,OWL.TransitiveProperty) not in g,name+' remains direct and non-transitive')
ck((N.containsPlace,RDF.type,OWL.TransitiveProperty) in g,'containsPlace transitive')
ck(not list(g.subjects(RDF.type,OWL.NamedIndividual)),'No examples written into deliverable')
for term in set(g.predicates())|set(g.objects(None,OWL.onProperty))|set(g.objects(None,OWL.onClass))|set(g.objects(None,RDFS.domain))|set(g.objects(None,RDFS.range)):
    if isinstance(term,URIRef) and str(term).startswith(str(N)):
        ck(any(g.triples((term,RDF.type,None))),'Local reference declared '+str(term).split('#')[-1])

test=Graph()
for t in g:test.add(t)
def put(s,p,o):test.add((N[s],p if isinstance(p,URIRef) else N[p],o))
put('testCampus',RDF.type,N.Campus)
put('testLibrary',RDF.type,N.LibraryBuilding)
put('testLibrary','locatedOnCampus',N.testCampus)
put('testLibrary','buildingName',Literal('Central Library'))
put('testLibrary','managedBy',N.testLibraryService)
put('testLibraryService',RDF.type,N.Organisation)
put('testLibrary','placeContactEmail',Literal('library@example.org'))
put('testLibrary','libraryServiceURL',Literal('https://example.org/library',datatype=X.anyURI))
put('testLibrary','hasLift',Literal(True))
put('testReading',RDF.type,N.ReadingRoom)
put('testReading','locatedInBuilding',N.testLibrary)
put('testReading','roomID',Literal('LIB-101'))
put('testReading','roomAreaSquareMetres',Literal('35.0',datatype=X.decimal))
put('testResidence',RDF.type,N.StudentResidence)
put('testResidence','residentCapacity',Literal(200,datatype=X.nonNegativeInteger))
put('testResidence','locatedOnCampus',N.testCampus)
put('testCourt',RDF.type,N.OutdoorCourt)
put('testCourt','locatedOnCampus',N.testCampus)
put('testCourt','isBookable',Literal(True))
put('testCourt','supportedSport',Literal('Basketball'))
put('testHall',RDF.type,N.SportsHall)
put('testHall','locatedOnCampus',N.testCampus)
put('testHall','supportedSport',Literal('Badminton'))
put('testComputers',RDF.type,N.EquipmentType)
put('testComputers','equipmentName',Literal('Computer'))
for name,count,booking in [('testStudyA',5,True),('testStudyB',6,False),('testStudyC',1,None),('testStudyZero',0,True)]:
    put(name,RDF.type,N.StudySpace)
    put('testLibrary','hasStudySpace',N[name])
    put(name,'hasEquipmentProvision',N[name+'Computers'])
    put(name+'Computers','equipmentType',N.testComputers)
    put(name+'Computers','quantity',Literal(count,datatype=X.nonNegativeInteger))
    if booking is not None:put(name,'isBookable',Literal(booking))
DeductiveClosure(OWLRL_Semantics,datatype_axioms=True).expand(test)
for t in [
    (N.testCampus,N.hasCampusPlace,N.testLibrary),
    (N.testCampus,N.containsPlace,N.testReading),
    (N.testLibrary,N.hasRoom,N.testReading),
    (N.testReading,N.withinCampus,N.testCampus),
    (N.testReading,RDF.type,N.LibraryReadingRoom),
    (N.testLibrary,N.placeName,Literal('Central Library')),
    (N.testReading,N.placeCode,Literal('LIB-101')),
    (N.testReading,N.areaSquareMetres,Literal('35.0',datatype=X.decimal)),
    (N.testStudyAComputers,N.provisionForRoom,N.testStudyA),
    (N.testStudyA,RDF.type,N.LibraryRoom),
    (N.testLibraryService,RDF.type,N.Organisation),
]:ck(t in test,'Inference '+str(t))
for wrong in [N.Person,N.Room]:ck((N.testLibrary,RDF.type,wrong) not in test,'Library not inferred as '+str(wrong).split('#')[-1])
for wrong in [N.Room,N.StudySpace]:ck((N.testResidence,RDF.type,wrong) not in test,'Residence capacity does not imply '+str(wrong).split('#')[-1])
for wrong in [N.Building,N.Room,N.SportsRoom]:ck((N.testCourt,RDF.type,wrong) not in test,'Outdoor court not inferred as '+str(wrong).split('#')[-1])
for wrong in [N.OutdoorSpace,N.Room]:ck((N.testHall,RDF.type,wrong) not in test,'Supported sport union does not imply '+str(wrong).split('#')[-1])
ck(not list(test.objects(N.testReading,N.isAccessible)),'Lift does not imply room accessibility')
ck((N.testStudyAComputers,N.provisionForRoom,N.testStudyB) not in test,'Equipment counts remain attached to their own room')
rows=list(test.query('''PREFIX ns: <https://example.org/hinton/newstudent#>
SELECT DISTINCT ?room ?bookable WHERE {
 ns:testLibrary ns:hasRoom ?room .
 ?room a ns:StudySpace; ns:hasEquipmentProvision ?record .
 ?record ns:equipmentType ns:testComputers; ns:quantity ?count .
 FILTER (?count > 0)
 OPTIONAL { ?room ns:isBookable ?bookable }
}'''))
ck({r.room for r in rows}=={N.testStudyA,N.testStudyB,N.testStudyC},'Computer room query includes bookable, unbookable and unknown; excludes zero computers')
ck({None if r.bookable is None else r.bookable.toPython() for r in rows}=={True,False,None},'Booking uncertainty preserved')
ck(not list(test.subjects(RDF.type,OWL.Nothing)),'No owl:Nothing inference in positive examples')
errors=list(test.objects(None,ERR.error));ck(not errors,'No OWL RL errors '+repr(errors))
report={'result':'passed','checks':len(checks),'named_entity_counts':counts,'source_definitions_unchanged':len(audit['unchanged_existing_definitions']),'inference':'RDFLib parsing and OWL RL closure over in-memory location, equipment and property-inheritance examples','full_owl2_dl_reasoner_run':False,'checks_passed':checks}
(WORK/'validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
print(json.dumps({k:v for k,v in report.items() if k!='checks_passed'},ensure_ascii=False,indent=2))
