from pathlib import Path
from copy import deepcopy
import hashlib
import json
import re
from lxml import etree as E

WORK = Path('/private/tmp/hinton-place-properties')
SOURCE = Path('/Users/peter/Desktop/campus_academic_ontology.owl')
OUTPUT = WORK / 'campus_academic_ontology.owl'
BASE = 'https://example.org/hinton/newstudent#'
NS = {
    'owl': 'http://www.w3.org/2002/07/owl#',
    'rdf': 'http://www.w3.org/1999/02/22-rdf-syntax-ns#',
    'rdfs': 'http://www.w3.org/2000/01/rdf-schema#',
    'xsd': 'http://www.w3.org/2001/XMLSchema#',
    'dcterms': 'http://purl.org/dc/terms/',
    'xml': 'http://www.w3.org/XML/1998/namespace',
}

def q(tag):
    a,b=tag.split(':',1)
    return '{'+NS[a]+'}'+b

def iri(name):
    if ':' in name:
        a,b=name.split(':',1)
        return NS[a]+b
    return BASE+name

def add(parent,tag,text=None,resource=None,**kwargs):
    node=E.SubElement(parent,q(tag))
    if text is not None: node.text=str(text)
    if resource is not None: node.set(q('rdf:resource'),iri(resource))
    for k,v in kwargs.items():node.set(q(k.replace('_',':')),v)
    return node

original=SOURCE.read_bytes()
(WORK/'original.owl').write_bytes(original)
root=E.fromstring(original,E.XMLParser(remove_blank_text=True))
before=deepcopy(root)
index={n.get(q('rdf:about'))[len(BASE):]:n for n in root if n.get(q('rdf:about'),'').startswith(BASE)}
changed=set();created=set()

def entity(name,tag):
    if name not in index:
        index[name]=E.SubElement(root,q(tag))
        index[name].set(q('rdf:about'),iri(name));created.add(name)
    assert index[name].tag==q(tag),name
    changed.add(name)
    return index[name]

def ann(node,tag,text):
    for old in node.findall(tag,NS):
        if old.get(q('xml:lang'))=='en':node.remove(old)
    add(node,tag,text,xml_lang='en')

def resource(node,tag,name):
    for old in node.findall(tag,NS):node.remove(old)
    add(node,tag,resource=name)

def domain(node,names,tag='rdfs:domain'):
    if isinstance(names,str):names=[names]
    for old in node.findall(tag,NS):node.remove(old)
    if len(names)==1:add(node,tag,resource=names[0])
    else:
        union=add(add(add(node,tag),'owl:Class'),'owl:unionOf',rdf_parseType='Collection')
        for name in names:add(union,'rdf:Description',rdf_about=iri(name))

def subproperty(name,parent):
    node=entity(name,'owl:DatatypeProperty' if index[name].tag==q('owl:DatatypeProperty') else 'owl:ObjectProperty')
    if not any(c.get(q('rdf:resource'))==iri(parent) for c in node.findall('rdfs:subPropertyOf',NS)):
        add(node,'rdfs:subPropertyOf',resource=parent)

def restriction(name,prop,target,kind,count,data=False):
    node=entity(name,'owl:Class')
    wrapper=add(node,'rdfs:subClassOf');r=add(wrapper,'owl:Restriction')
    add(r,'owl:onProperty',resource=prop)
    add(r,'owl:'+kind,count,rdf_datatype=iri('xsd:nonNegativeInteger'))
    add(r,'owl:onDataRange' if data else 'owl:onClass',resource=target)

support={
    'EquipmentType': ('Equipment type','A controlled equipment category, such as computer or projector. A category is represented by an individual of this class, not by a count or a physical room.'),
    'EquipmentProvision': ('Equipment provision','A current record linking exactly one room to exactly one equipment category and its non-negative quantity. Use separate records for different rooms or equipment categories; changes over time require separate snapshots or a later temporal extension.'),
}
for name,(label,comment) in support.items():
    node=entity(name,'owl:Class');resource(node,'rdfs:subClassOf','owl:Thing')
    ann(node,'rdfs:label',label);ann(node,'rdfs:comment',comment)

# domain, range, English label and comment, Chinese dictionary explanation.
object_new={
    'containsPlace': ('Place','Place','Contains place','The inverse of transitive spatiallyPartOf. Contains a place directly or indirectly; do not count these values as only direct children.','直接或间接包含的地点；spatiallyPartOf 的逆属性，具有传递性。'),
    'hasRoom': ('Building','Room','Has room','The inverse of locatedInBuilding: an individually identifiable room directly contained in a building.','建筑直接包含的房间；与 locatedInBuilding 互为逆属性。'),
    'hasCampusPlace': ('Campus',['Building','OutdoorSpace'],'Has campus place','The inverse of locatedOnCampus. Relates a campus to its buildings or outdoor spaces; this does not include rooms as direct campus children.','校区直接包含的建筑或室外空间；与 locatedOnCampus 互为逆属性。'),
    'managedBy': ('Place','Organisation','Managed by','An organisation responsible for operating or managing a place. Multiple managing organisations are allowed. This is distinct from spatial containment.','地点的管理或运营组织，允许多个；不能用来表示地点在某建筑内部。'),
    'hasEquipmentProvision': ('Room','EquipmentProvision','Has equipment provision','A room-specific equipment category and quantity record. Counts for different equipment types are kept in separate records.','房间关联设备配置记录；每种设备的数量通过各自记录保存。'),
    'provisionForRoom': ('EquipmentProvision','Room','Provision for room','The room described by an equipment provision record; inverse of hasEquipmentProvision.','配置记录所描述的唯一房间；hasEquipmentProvision 的逆属性。'),
    'equipmentType': ('EquipmentProvision','EquipmentType','Equipment type','The single equipment category counted by a provision record.','配置记录统计的唯一设备类型，例如 ComputerEquipmentType。'),
}
for name,(d,r,label,comment,zh) in object_new.items():
    node=entity(name,'owl:ObjectProperty');domain(node,d);domain(node,r,'rdfs:range')
    ann(node,'rdfs:label',label);ann(node,'rdfs:comment',comment)
for name,target in [('containsPlace','spatiallyPartOf'),('hasRoom','locatedInBuilding'),('hasCampusPlace','locatedOnCampus'),('provisionForRoom','hasEquipmentProvision')]:
    resource(index[name],'owl:inverseOf',target)
add(index['containsPlace'],'rdf:type',resource='owl:TransitiveProperty')
for name in ['hasRoom','hasCampusPlace']:subproperty(name,'containsPlace')
for name in ['hasLibraryRoom','hasStudySpace']:subproperty(name,'hasRoom')
subproperty('withinCampus','spatiallyPartOf')

data_new={
    'placeName': ('Place','string','Place name','A human-readable name of a place. Names do not identify places uniquely.','地点通用名称；子类专用名称可自动汇总到此属性。'),
    'placeCode': ('Place','string','Place code','A code for a place within its documented naming scope. Codes need not be globally unique across campuses or kinds of place.','地点通用编码；作用范围需说明，不声明为全校唯一键。'),
    'locationDescription': ('Place','string','Location description','A textual wayfinding description such as the entrance to use or the side of a square. It does not replace formal location relationships.','寻路说明，例如入口方向、相邻标志物；不能替代正式位置关系。'),
    'placeContactEmail': ('Place','string','Place contact email','A public service contact email for a place. Do not reuse the Person-domain emailAddress property for places.','地点服务邮箱；避免复用 Person 的 emailAddress 而误推断地点是人员。'),
    'placeContactTelephone': ('Place','string','Place contact telephone','A public service telephone number for a place; preserve prefixes and leading zeroes.','地点服务电话，以字符串保留区号和前导零。'),
    'accessPolicy': ('Place','string','Access policy','Published access conditions such as card access, opening restrictions or safety requirements. This text is not an automated eligibility rule.','门禁、开放对象、入场要求等说明；文字本身不执行资格判定。'),
    'serviceURL': ('Place','anyURI','Service URL','An official service information page for the place, distinct from its map or booking page.','官方服务信息页；与地图和预约入口分别记录。'),
    'hasFixedSeating': ('TeachingRoom','boolean','Has fixed seating','Whether the teaching room has fixed rather than freely movable seating in the current recorded configuration. Missing means unknown.','教室是否采用固定座椅；缺失为未知。'),
    'requiresSafetyInduction': ('Laboratory','boolean','Requires safety induction','Whether users must complete a safety induction before using this laboratory. Missing means unknown; true does not mean a given student has completed the induction.','使用实验室前是否要求安全培训；不能据此推断某学生已完成培训。'),
    'collectionDescription': ('BookStackRoom','string','Collection description','The subject scope or categories of resources held in a book stack room. It is not an item-level catalogue.','书库馆藏主题或类别说明；具体书目使用图书馆目录。'),
    'supportedSport': (['SportsBuilding','SportsRoom','OutdoorSpace'],'string','Supported sport','A supported sport or activity name. Multiple values may be recorded using a shared controlled vocabulary. It does not imply equipment availability or current booking availability.','支持的运动，可多值，例如 Basketball、Badminton；不代表已配齐设备或当前可预约。'),
    'hasFloodlighting': ('OutdoorSpace','boolean','Has floodlighting','Whether an outdoor space has floodlighting. This does not imply that it is open at night. Missing means unknown.','室外空间是否配有照明设施；不能据此推出夜间开放。'),
    'residentCapacity': ('StudentResidence','nonNegativeInteger','Resident capacity','The current permitted number of residents in a student residence building. This is not the Room-domain capacity property.','宿舍楼核定住宿人数；与 Room 的 capacity 分开，避免把宿舍楼推断为房间。'),
    'cateringInformation': ('DiningBuilding','string','Catering information','Published meal types or dietary-service information. Link to current details using serviceURL.','餐饮服务或饮食选择说明；时效性详情可由 serviceURL 指向官方页面。'),
    'equipmentName': ('EquipmentType','string','Equipment name','The display name of an equipment category.','设备类别名称，例如电脑、投影仪。'),
    'quantity': ('EquipmentProvision','nonNegativeInteger','Quantity','The non-negative number of items of the category held in the room for this provision record. A zero count does not qualify as available equipment.','该房间该类设备的数量；查询有电脑时要求数量大于零。'),
}
for name,(d,t,label,comment,zh) in data_new.items():
    node=entity(name,'owl:DatatypeProperty');domain(node,d);resource(node,'rdfs:range','xsd:'+t)
    ann(node,'rdfs:label',label);ann(node,'rdfs:comment',comment)

# Generalise only genuinely shared Place properties. Retain the old specialised
# names as subproperties so existing assertions and downstream queries survive.
broaden={
    'postalAddress': 'The public postal address of a place where applicable. A room can share its containing building address; do not infer an address automatically from location alone.',
    'mapURL': 'An official map link for a place, including campuses, buildings, rooms or outdoor spaces.',
    'latitude': 'A representative latitude of a place, in degrees from -90 to 90. A room-specific coordinate is optional and is not automatically inherited from its building.',
    'longitude': 'A representative longitude of a place, in degrees from -180 to 180. A room-specific coordinate is optional and is not automatically inherited from its building.',
}
for name,comment in broaden.items():
    node=entity(name,'owl:DatatypeProperty');domain(node,'Place');ann(node,'rdfs:comment',comment)
domain(entity('surfaceType','owl:DatatypeProperty'),['Room','OutdoorSpace'])
ann(index['surfaceType'],'rdfs:comment','The surface material of a room or outdoor space, such as wood, grass or a synthetic sports surface. The domain is one union expression.')

aliases={
    'campusName':'placeName','buildingName':'placeName','roomName':'placeName','spaceName':'placeName',
    'buildingCode':'placeCode','roomID':'placeCode','spaceCode':'placeCode',
    'campusAddress':'postalAddress','campusMapURL':'mapURL','libraryServiceURL':'serviceURL',
    'openingHours':'openingHoursText','roomAreaSquareMetres':'areaSquareMetres',
}
for name,parent in aliases.items():subproperty(name,parent)

restriction('EquipmentProvision','equipmentType','EquipmentType','qualifiedCardinality',1)
restriction('EquipmentProvision','provisionForRoom','Room','qualifiedCardinality',1)
restriction('EquipmentProvision','quantity','xsd:nonNegativeInteger','qualifiedCardinality',1,data=True)
for cls,prop,datatype in [('TeachingRoom','hasFixedSeating','boolean'),('Laboratory','requiresSafetyInduction','boolean'),('OutdoorSpace','hasFloodlighting','boolean'),('StudentResidence','residentCapacity','nonNegativeInteger')]:
    restriction(cls,prop,'xsd:'+datatype,'maxQualifiedCardinality',1,data=True)

ontology=root.find('owl:Ontology',NS)
ontology.find('owl:versionInfo',NS).text='2.8'
for c in ontology.findall('rdfs:comment',NS):
    if 'Activity scheduling, equipment provision and facility matching are deferred.' in (c.text or ''):
        c.text=c.text.replace('Activity scheduling, equipment provision and facility matching are deferred.','Room equipment provision is included. Activity scheduling and facility matching remain deferred.')
    if 'The 30 additional data properties adopted in version 2.4' in (c.text or ''):
        c.text='The 30 additional data properties adopted in version 2.4 are retained. Version 2.8 adds shared Place properties, inverse location relationships and room equipment provision. There are 62 named classes, 30 named object properties and 106 named data properties. English labels and comments, identifiers and existing non-spatial reasoning axioms are retained.'
add(ontology,'dcterms:source','Place property design, 28 October 2026, based on the version 1.3 class design and the room equipment questions Q8 and Q9.',xml_lang='en')
add(ontology,'rdfs:comment','Shared location properties are defined on Place and inherited by its subclasses. Existing specialised property names are retained as subproperties where their meanings align. Supported sports use a shared string vocabulary; booking and opening-hour fields describe policy, not live availability. The equipment support classes are not subclasses of Place. No illustrative individuals are included.',xml_lang='en')

untouched=[]
for old in before:
    about=old.get(q('rdf:about'),'');local=about[len(BASE):] if about.startswith(BASE) else None
    if old.tag==q('owl:Ontology') or local in changed:continue
    assert any(E.tostring(old,method='c14n')==E.tostring(n,method='c14n') for n in root),about or old.tag
    untouched.append(about or old.tag)

E.indent(root,space='  ')
OUTPUT.write_bytes(E.tostring(root,encoding='utf-8',xml_declaration=True,pretty_print=True).replace(b'\n',b'\r\n'))
audit={
    'source':str(SOURCE),'source_sha256':hashlib.sha256(original).hexdigest(),
    'output_sha256':hashlib.sha256(OUTPUT.read_bytes()).hexdigest(),
    'new_classes':sorted(support),'new_object_properties':sorted(object_new),'new_data_properties':sorted(data_new),
    'modified_existing_entities':sorted(changed-created),'unchanged_existing_definitions':untouched,
    'aliases':aliases,'broadened_properties':list(broaden)+['surfaceType'],
    'object_new':object_new,'data_new':data_new,
}
(WORK/'audit.json').write_text(json.dumps(audit,ensure_ascii=False,indent=2))
print(json.dumps({k:audit[k] for k in ['new_classes','new_object_properties','new_data_properties','broadened_properties']},ensure_ascii=False,indent=2))
print('Unchanged existing definitions:',len(untouched))
