"""Simple XBRL parser implementation."""

import re
import urllib.parse
import xml.etree.ElementTree as ET
import xml.sax.saxutils as saxutils
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Set, Tuple

from .file_loader import FileLoader


def resolve_url(base_url: str, relative_url: str) -> str:
    """Resolve a potentially relative URL against a base path or URL."""

    # Already absolute URL
    if relative_url.startswith(("http://", "https://")):
        return relative_url

    # Handle file system paths (base URL is a local path)
    if not base_url.startswith(("http://", "https://")):
        base_path = Path(base_url).parent
        resolved_path = (base_path / relative_url).resolve()
        return str(resolved_path)

    # Handle HTTP URLs - ensure base_url ends with / for proper resolution
    if not base_url.endswith("/"):
        base_url = base_url.rsplit("/", 1)[0] + "/"
    return urllib.parse.urljoin(base_url, relative_url)


# ====================
# Model Data Classes
# ====================


@dataclass
class QName:
    """Qualified name for XBRL elements."""

    local_name: str  # Local part (e.g., "CashAndDeposits")
    namespace_uri: Optional[str] = None  # Full namespace URI
    prefix: Optional[str] = None  # Namespace prefix (e.g., "jppfs_cor")

    @property
    def full_name(self) -> str:
        """Get the full qualified name with prefix."""
        if self.prefix:
            return f"{self.prefix}:{self.local_name}"
        return self.local_name

    @classmethod
    def parse(cls, name: str, namespaces: Optional[Dict[str, str]] = None) -> "QName":
        """
        Parse a qualified name string into QName components.

        Args:
            name: Qualified name string (e.g., "jppfs_cor:CashAndDeposits")
            namespaces: Optional namespace mapping dict

        Returns:
            QName instance
        """
        if ":" in name:
            prefix, local_name = name.split(":", 1)
            namespace_uri = namespaces.get(prefix) if namespaces else None
            return cls(
                local_name=local_name, prefix=prefix, namespace_uri=namespace_uri
            )
        return cls(local_name=name)

    def __str__(self) -> str:
        """String representation returns the full qualified name."""
        return self.full_name


def elem_to_simple_tokens(elem: ET.Element) -> List[str]:
    """
    Convert an XML element to a list of simple text tokens.

    ET.tostring() includes prefixes and namespace declarations, which we want to avoid.
    """
    # Output tag token of html only. If tag is not html, output text only.
    is_html = elem.tag.startswith("{http://www.w3.org/1999/xhtml}")
    local_name = elem.tag.split("}", 1)[-1] if "}" in elem.tag else elem.tag
    attrs = "".join(f' {k}="{v}"' for k, v in elem.attrib.items())
    # Check if element is self-closing (no text content and no children)
    if (
        not elem.text
        and len(elem) == 0
        and local_name in ("br", "hr", "img", "input", "meta", "link", "col")
    ):
        # Self-closing tag
        tokens = [f"<{local_name}{attrs}/>"] if is_html else []
        # Add tail if present
        if elem.tail:
            tokens.append(saxutils.escape(elem.tail))
        return tokens
    tokens = [f"<{local_name}{attrs}>"] if is_html else []
    if elem.text:
        tokens.append(saxutils.escape(elem.text))
    for child in elem:
        tokens.extend(elem_to_simple_tokens(child))
    if is_html:
        tokens.append(f"</{local_name}>")
    if elem.tail:
        tokens.append(saxutils.escape(elem.tail))
    return tokens


@dataclass
class Fact:
    """Single XBRL fact - kept as simple strings."""

    qname: QName  # Parsed QName object
    raw_element: ET.Element  # Original XML element to extract text data.
    context_ref: str
    is_numeric: bool
    is_nil: bool
    escape: bool
    unit_ref: Optional[str] = None
    scale: Optional[str] = None
    sign: Optional[str] = None
    format: Optional[str] = None
    attrs: Dict[str, str] = None

    @property
    def value(self) -> Optional[str | Decimal]:
        """Get the concatenated text value of the fact."""
        if self.is_nil:
            return None
        if self.is_numeric:
            # For numeric facts, parse as Decimal for precise financial calculations
            text = self.raw_element.text.strip()
            try:
                # Check if this is a Japanese currency format (ixt:numunitdecimal)
                if self.format == "ixt:numunitdecimal":
                    # Parse Japanese yen format like "800円0銭" -> 800.00
                    match = re.match(r"(\d+(?:,\d{3})*)円(\d+)銭", text)
                    if match:
                        yen = match.group(1).replace(",", "")
                        sen = match.group(2)
                        # Convert to decimal (1 yen = 100 sen)
                        value = Decimal(yen) + Decimal(sen) / 100
                    else:
                        # Fallback: try to extract just numbers
                        value = Decimal(
                            text.replace(",", "").replace("円", "").replace("銭", "")
                        )
                else:
                    # Remove commas for thousands separator
                    value = Decimal(text.replace(",", ""))

                if self.scale:
                    # Apply scale factor (e.g., scale=6 means multiply by 10^6)
                    value *= Decimal(10) ** int(self.scale)
                if self.sign == "-":
                    value = -value
                return value
            except (ValueError, TypeError):
                pass  # Return as-is if conversion fails
        if self.escape:
            tokens = (
                [saxutils.escape(self.raw_element.text)]
                if self.raw_element.text
                else []
            )
            for child in self.raw_element:
                tokens.extend(elem_to_simple_tokens(child))
        else:
            tokens = list(self.raw_element.itertext())
        return "".join(tokens).strip() if tokens else None


@dataclass
class Arc:
    from_qname: QName  # QName from linkbase
    to_qname: QName  # QName from linkbase
    role: str  # Role URI (e.g., http://example.com/role/BalanceSheet)
    arcrole: str  # Arc role (e.g., http://www.xbrl.org/2003/arcrole/parent-child for presentation, summation-item for calculation)
    weight: Optional[Decimal] = (
        None  # Weight for calculation arcs (e.g., 1.0 for addition, -1.0 for subtraction)
    )


@dataclass
class Label:
    """Label from label linkbase."""

    qname: QName  # QName from label linkbase (from link:loc href)
    text: str  # The label text content
    lang: str  # Language code (e.g., "ja", "en") from xml:lang
    link_role: str  # Role from link:labelLink xlink:role (e.g., "http://www.xbrl.org/2003/role/link")
    label_role: str  # Label role from link:label xlink:role (e.g., "http://www.xbrl.org/2003/role/label")
    arcrole: str  # Arc role from link:labelArc xlink:arcrole (typically "http://www.xbrl.org/2003/arcrole/concept-label")
    priority: Optional[Decimal]  # Optional priority attribute on label arcs


@dataclass(frozen=True)
class XsdSchema:
    """Internal class to hold XSD schema information."""

    target_namespace: str  # The targetNamespace of the XSD schema
    prefix: Optional[str]  # The namespace prefix for this namespace (if any)
    elements: Dict[str, str]  # Map from element ID to element name
    linkbase_refs: Dict[
        str, List[Tuple[str, str]]
    ]  # Map from linkbase role type to list of (href, role) tuples
    imports: List[str]  # List of imported schema URLs from xsd:import elements


# ====================
# XmlParser Classes
# ====================


class XmlParser:
    """Base interface for XML event parsers."""

    def on_xml_event(self, event: str, elem) -> None:
        """
        Process XML parsing events.

        Args:
            event: Event type from ET.iterparse
            elem: XML element or namespace tuple
        """
        raise NotImplementedError


class XsdPathParser(XmlParser):
    """Parser for extracting XSD schema file paths from inline XBRL files."""

    def __init__(self, base_path: Path):
        """
        Initialize the XSD path parser.

        Args:
            base_path: Base path of the current iXBRL file for resolving relative paths
        """
        self.base_path = base_path
        self.schema_ref_tag = None
        self.href_attr_name = None
        self.xsd_paths: List[Path] = []

    def on_xml_event(self, event: str, elem) -> None:
        """
        Process XML parsing events to extract XSD schema references.

        Args:
            event: Event type from ET.iterparse
            elem: XML element or namespace tuple
        """
        if event == "start-ns":
            prefix, uri = elem
            if prefix == "link":
                self.schema_ref_tag = f"{{{uri}}}schemaRef"
            elif prefix == "xlink":
                self.href_attr_name = f"{{{uri}}}href"

        elif event == "end" and self.schema_ref_tag and self.href_attr_name:
            # Look for link:schemaRef element
            if elem.tag == self.schema_ref_tag:
                href = elem.get(self.href_attr_name)
                if href:
                    # Resolve relative to iXBRL file location
                    xsd_path = self.base_path.parent / href
                    if xsd_path.exists() and xsd_path not in self.xsd_paths:
                        self.xsd_paths.append(xsd_path)


class LabelResourceParser(XmlParser):
    """Parser for extracting label resources from label linkbase files."""

    LABEL_LOC_ROLE = "http://www.xbrl.org/2003/role/label"

    def __init__(self, file_loader: FileLoader, base_url: str):
        """Initialize the label resource parser."""
        if file_loader is None:
            raise ValueError("LabelResourceParser requires a FileLoader instance")

        self.file_loader = file_loader
        self.base_url = str(base_url)
        self.label_resources: Dict[str, List[Tuple[str, str, str]]] = {}
        self.remote_label_resources: Dict[
            str, Dict[str, List[Tuple[str, str, str]]]
        ] = {}
        self.label_tag = None
        self.loc_tag = None
        self.type_attr_name = None
        self.label_attr_name = None
        self.role_attr_name = None
        self.href_attr_name = None
        self.lang_attr_name = "{http://www.w3.org/XML/1998/namespace}lang"

    def on_xml_event(self, event: str, elem) -> None:
        """
        Process XML parsing events to extract label resources.

        Args:
            event: Event type from ET.iterparse
            elem: XML element or namespace tuple
        """
        if event == "start-ns":
            prefix, uri = elem
            if prefix == "link":
                self.label_tag = f"{{{uri}}}label"
                self.loc_tag = f"{{{uri}}}loc"
            elif prefix == "xlink":
                self.type_attr_name = f"{{{uri}}}type"
                self.label_attr_name = f"{{{uri}}}label"
                self.role_attr_name = f"{{{uri}}}role"
                self.href_attr_name = f"{{{uri}}}href"
            elif prefix == "xml":
                self.lang_attr_name = f"{{{uri}}}lang"

        elif event == "end":
            if self.label_tag and elem.tag == self.label_tag:
                self._process_label_element(elem)
            elif self.loc_tag and elem.tag == self.loc_tag:
                self._process_loc_element(elem)

    def _process_label_element(self, elem) -> None:
        """Handle label resource elements within a label linkbase."""
        xlink_type = elem.get(self.type_attr_name) if self.type_attr_name else None

        if xlink_type == "resource":
            label_label = elem.get(self.label_attr_name)
            label_role = elem.get(self.role_attr_name)
            label_lang = elem.get(self.lang_attr_name)
            label_text = elem.text or ""

            if label_label:
                if label_label not in self.label_resources:
                    self.label_resources[label_label] = []
                self.label_resources[label_label].append(
                    (label_text, label_lang or "", label_role or "")
                )

    def _process_loc_element(self, elem) -> None:
        """Handle locator elements that point to remote label resources."""
        if not self.type_attr_name or not self.label_attr_name:
            return

        xlink_type = elem.get(self.type_attr_name)
        if xlink_type != "locator":
            return

        loc_role = elem.get(self.role_attr_name) if self.role_attr_name else None
        if loc_role != self.LABEL_LOC_ROLE:
            return

        loc_href = elem.get(self.href_attr_name) if self.href_attr_name else None
        loc_label = elem.get(self.label_attr_name)

        if not loc_href or not loc_label:
            return

        resource_href, fragment = self._split_href(loc_href)
        if not fragment:
            return

        resolved_url = resolve_url(self.base_url, resource_href)

        remote_labels = self._get_or_load_remote_label_map(resolved_url)
        if not remote_labels:
            return

        entries = remote_labels.get(fragment)
        if not entries:
            return

        if loc_label not in self.label_resources:
            self.label_resources[loc_label] = []
        self.label_resources[loc_label].extend(entries)

    def _split_href(self, href: str) -> Tuple[str, Optional[str]]:
        """Split an href into base URL and fragment identifier."""
        if "#" in href:
            base, fragment = href.split("#", 1)
            return base, fragment
        return href, None

    def _get_or_load_remote_label_map(
        self, resolved_url: str
    ) -> Dict[str, List[Tuple[str, str, str]]]:
        """Load and cache remote label resources from the given URL."""
        if resolved_url in self.remote_label_resources:
            return self.remote_label_resources[resolved_url]

        if not self.file_loader:
            self.remote_label_resources[resolved_url] = {}
            return {}

        label_path = self.file_loader.fetch(resolved_url)
        if label_path is None:
            self.remote_label_resources[resolved_url] = {}
            return {}

        label_map = self._parse_remote_label_file(label_path)
        self.remote_label_resources[resolved_url] = label_map
        return label_map

    def _parse_remote_label_file(
        self, label_path: Path
    ) -> Dict[str, List[Tuple[str, str, str]]]:
        """Parse a remote label definition file and return label entries."""
        remote_map: Dict[str, List[Tuple[str, str, str]]] = {}
        label_tag = None
        type_attr_name = None
        label_attr_name = None
        role_attr_name = None
        lang_attr_name = "{http://www.w3.org/XML/1998/namespace}lang"

        try:
            for event, elem in ET.iterparse(
                str(label_path), events=["start-ns", "end"]
            ):
                if event == "start-ns":
                    prefix, uri = elem
                    if prefix == "link":
                        label_tag = f"{{{uri}}}label"
                    elif prefix == "xlink":
                        type_attr_name = f"{{{uri}}}type"
                        label_attr_name = f"{{{uri}}}label"
                        role_attr_name = f"{{{uri}}}role"
                    elif prefix == "xml":
                        lang_attr_name = f"{{{uri}}}lang"
                elif event == "end" and label_tag and elem.tag == label_tag:
                    xlink_type = elem.get(type_attr_name) if type_attr_name else None

                    if xlink_type != "resource":
                        elem.clear()
                        continue

                    label_id = elem.get("id")
                    if not label_id and label_attr_name:
                        label_id = elem.get(label_attr_name)

                    if not label_id:
                        elem.clear()
                        continue

                    label_role = elem.get(role_attr_name) if role_attr_name else None
                    label_lang = elem.get(lang_attr_name)
                    label_text = elem.text or ""

                    if label_id not in remote_map:
                        remote_map[label_id] = []
                    remote_map[label_id].append(
                        (label_text, label_lang or "", label_role or "")
                    )
                    elem.clear()
        except Exception:
            return {}

        return remote_map


# ====================
# Parser Class
# ====================


class Parser:
    """Lightweight XBRL parser for Japanese financial data."""

    def __init__(
        self,
        file_loader: Optional[FileLoader] = None,
        shared_xsd_cache: Optional[Dict[str, XsdSchema]] = None,
        follow_xsd_imports: bool = True,
    ):
        """
        Initialize the parser.

        Args:
            file_loader: Optional FileLoader instance for fetching remote files.
                        If not provided, creates a default FileLoader.
            shared_xsd_cache: Optional shared cache for remote XSD schemas.
                             Can be shared between Parser instances to avoid redundant
                             network requests for the same remote schemas.
            follow_xsd_imports: Whether to recursively follow XSD import statements
                               to collect linkbase refs from all imported schemas.
                               Default is True for more accurate and complete parsing.
        """
        self.file_loader = file_loader or FileLoader()
        self.ixbrl_files: List[Path] = []
        self.root_xsd_files: List[Path] = (
            []
        )  # Root XSD schema files referenced by iXBRL
        self.linkbase_refs: Dict[str, List[Tuple[str, str]]] = (
            {}
        )  # role -> [(href, arcrole)]
        # Cache for XSD schemas by URL - instance-specific cache
        self.xsd_schemas_by_url: Dict[str, XsdSchema] = {}
        # Optional shared cache for remote XSD schemas
        self.shared_xsd_cache = shared_xsd_cache
        # Whether to follow XSD imports
        self.follow_xsd_imports = follow_xsd_imports

    def prepare_ixbrl(self, manifest_path: Path) -> None:
        """
        Parse manifest file and prepare list of inline XBRL files.

        Args:
            manifest_path: Path to manifest XML file
        """
        manifest_path = Path(manifest_path)
        if not manifest_path.exists():
            raise FileNotFoundError(f"Manifest file not found: {manifest_path}")

        base_dir = manifest_path.parent
        self.ixbrl_files = []

        tree = ET.parse(manifest_path)
        root = tree.getroot()

        for elem in root.iter():
            if elem.tag.endswith("ixbrl"):
                if elem.text:
                    ixbrl_path = base_dir / elem.text
                    if ixbrl_path.exists():
                        self.ixbrl_files.append(ixbrl_path)

        if not self.ixbrl_files:
            raise ValueError(f"No inline XBRL files found in {base_dir}")

    def load_facts(self) -> Iterator[Fact]:
        """
        Load and yield facts from prepared inline XBRL files.

        Yields:
            Fact objects parsed from the inline XBRL files
        """
        if not self.ixbrl_files:
            raise RuntimeError(
                "No inline XBRL files prepared. Call prepare_ixbrl() first."
            )

        # Collect XSD paths locally first
        collected_xsd_paths: List[Path] = []

        for ixbrl_path in self.ixbrl_files:
            namespaces: Dict[str, str] = {}
            nonNumeric_tag = None
            nonFraction_tag = None
            nil_attr_name = None

            # Create XSD path parser for this file
            xsd_parser = XsdPathParser(ixbrl_path)

            # Use iterparse for streaming parsing
            for event, elem in ET.iterparse(
                str(ixbrl_path), events=["start-ns", "end"]
            ):
                # Let XSD parser process the event
                xsd_parser.on_xml_event(event, elem)

                if event == "start-ns":
                    # Capture namespace declarations
                    prefix, uri = elem

                    # Store namespace mapping
                    if prefix:
                        namespaces[prefix] = uri

                    if prefix == "xsi":
                        nil_attr_name = f"{{{uri}}}nil"
                    elif prefix == "ix":
                        # Build the full tag names for exact matching
                        nonNumeric_tag = f"{{{uri}}}nonNumeric"
                        nonFraction_tag = f"{{{uri}}}nonFraction"

                elif nonNumeric_tag and nonFraction_tag:
                    if event == "end":
                        # Process fact elements if we know the ix namespace
                        fact = None
                        if elem.tag == nonFraction_tag:
                            fact = self._extract_fact(
                                elem, nil_attr_name, namespaces, is_numeric=True
                            )
                        elif elem.tag == nonNumeric_tag:
                            fact = self._extract_fact(
                                elem, nil_attr_name, namespaces, is_numeric=False
                            )

                        if fact:
                            yield fact

            # Collect XSD paths from this file into local list
            for xsd_path in xsd_parser.xsd_paths:
                if xsd_path not in collected_xsd_paths:
                    collected_xsd_paths.append(xsd_path)

        # Only update self.root_xsd_files after successfully processing all files
        if not self.root_xsd_files:  # Only update if not already populated
            self.root_xsd_files = collected_xsd_paths

    def _extract_fact(
        self,
        elem: ET.Element,
        nil_attr_name: str,
        namespaces: Dict[str, str],
        is_numeric: bool,
    ) -> Fact:
        """
        Extract a fact from an inline XBRL element.

        Args:
            elem: XML element containing the fact

        Returns:
            Fact object or None if required attributes missing
        """
        # Get required attributes
        name = elem.get("name")
        context_ref = elem.get("contextRef")

        if not name or not context_ref:
            return None

        # Parse the QName
        qname = QName.parse(name, namespaces)

        # Check for xsi:nil attribute using cached attribute name
        is_nil = elem.get(nil_attr_name) == "true"

        return Fact(
            qname=qname,
            raw_element=elem,
            context_ref=context_ref,
            is_numeric=is_numeric,
            is_nil=is_nil,
            escape=elem.get("escape") == "true",
            unit_ref=elem.get("unitRef"),
            scale=elem.get("scale"),
            sign=elem.get("sign"),
            format=elem.get("format"),
            attrs=elem.attrib,
        )

    def find_xsd_files(self) -> List[Path]:
        """
        Find root XSD schema files directly referenced by the inline XBRL files.

        Returns:
            List of paths to root XSD files

        Note:
            If load_facts() has been called, this method will return the XSD files
            found during that parsing. Otherwise, it will parse the iXBRL files
            specifically to find XSD references.
        """
        if self.root_xsd_files:
            return self.root_xsd_files  # Already found

        # Parse iXBRL files to find XSD references
        for ixbrl_path in self.ixbrl_files:
            xsd_parser = XsdPathParser(ixbrl_path)

            for event, elem in ET.iterparse(
                str(ixbrl_path), events=["start-ns", "end"]
            ):
                xsd_parser.on_xml_event(event, elem)

            # Merge XSD paths found in this file
            for xsd_path in xsd_parser.xsd_paths:
                if xsd_path not in self.root_xsd_files:
                    self.root_xsd_files.append(xsd_path)

        if not self.root_xsd_files:
            raise ValueError(f"No XSD schema files found in {self.ixbrl_files}")

        return self.root_xsd_files

    def _get_or_parse_xsd_schema(self, resolved_schema_url: str) -> Optional[XsdSchema]:
        """
        Get XSD schema from cache or parse it if not cached.

        Args:
            resolved_schema_url: Resolved absolute URL or path to the XSD schema

        Returns:
            XsdSchema object if found or parsed successfully, None otherwise
        """
        # Check if it's a remote URL
        is_remote_url = resolved_schema_url.startswith(("http://", "https://"))

        # First check instance cache
        if resolved_schema_url in self.xsd_schemas_by_url:
            return self.xsd_schemas_by_url[resolved_schema_url]

        # Then check shared cache for remote URLs
        if (
            is_remote_url
            and self.shared_xsd_cache is not None
            and resolved_schema_url in self.shared_xsd_cache
        ):
            return self.shared_xsd_cache[resolved_schema_url]

        # Parse if not cached
        schema = self._parse_xsd_schema(resolved_schema_url)
        # Store in instance cache
        self.xsd_schemas_by_url[resolved_schema_url] = schema
        # Also store in shared cache if it's a remote URL
        if is_remote_url and self.shared_xsd_cache is not None:
            self.shared_xsd_cache[resolved_schema_url] = schema
        return schema

    def _parse_xsd_schema(self, xsd_url: str) -> XsdSchema:
        """
        Parse XSD schema to extract namespace and element definitions.

        Args:
            xsd_url: URL of the XSD schema file

        Returns:
            XsdSchema object containing parsed schema information
        """
        # Fetch the XSD file
        xsd_path = self.file_loader.fetch(xsd_url)
        if xsd_path is None:
            # Return empty schema if cannot fetch
            return XsdSchema(
                target_namespace=None,
                prefix=None,
                elements={},
                linkbase_refs={},
                imports=[],
            )
        # We'll extract these from the root element after parsing
        target_namespace = None
        schema_prefix = None
        elements = {}
        linkbase_refs = {}
        imports = []

        # Tags and attributes we need (will be set when namespace is discovered)
        xsd_schema_tag = None
        xsd_element_tag = None
        xsd_import_tag = None
        linkbase_ref_tag = None
        href_attr_name = None
        role_attr_name = None

        # Namespace prefix mapping
        namespace_map = {}

        for event, elem in ET.iterparse(
            str(xsd_path), events=["start-ns", "start", "end"]
        ):
            if event == "start-ns":
                prefix, uri = elem
                # Build namespace mapping
                namespace_map[prefix] = uri
                # Detect XSD namespace even when declared without explicit prefix
                if prefix == "xsd" or (
                    prefix == "" and uri == "http://www.w3.org/2001/XMLSchema"
                ):
                    xsd_schema_tag = f"{{{uri}}}schema"
                    xsd_element_tag = f"{{{uri}}}element"
                    xsd_import_tag = f"{{{uri}}}import"
                elif prefix == "link":
                    linkbase_ref_tag = f"{{{uri}}}linkbaseRef"
                elif prefix == "xlink":
                    href_attr_name = f"{{{uri}}}href"
                    role_attr_name = f"{{{uri}}}role"

            elif event == "start":
                # Get targetNamespace from root element
                if xsd_schema_tag and elem.tag == xsd_schema_tag:
                    target_namespace = elem.get("targetNamespace")

                    # Find the prefix for the target namespace
                    for prefix, uri in namespace_map.items():
                        if uri == target_namespace:
                            schema_prefix = prefix
                            break

            elif event == "end":
                # Process element definitions
                if xsd_element_tag and elem.tag == xsd_element_tag:
                    elem_id = elem.get("id")
                    elem_name = elem.get("name")
                    if elem_id and elem_name:
                        elements[elem_id] = elem_name
                    # Clear element to save memory
                    elem.clear()
                # Process import statements
                elif xsd_import_tag and elem.tag == xsd_import_tag:
                    schema_location = elem.get("schemaLocation")
                    if schema_location:
                        # Resolve relative to XSD URL
                        resolved_import_url = resolve_url(xsd_url, schema_location)
                        imports.append(resolved_import_url)
                    # Clear element to save memory
                    elem.clear()
                # Process linkbase references
                elif (
                    linkbase_ref_tag and href_attr_name and elem.tag == linkbase_ref_tag
                ):
                    href = elem.get(href_attr_name)
                    role = elem.get(role_attr_name)
                    if href and role:
                        # Extract role type (e.g., 'labelLinkbaseRef')
                        role_type = role.split("/")[-1] if "/" in role else role

                        if role_type not in linkbase_refs:
                            linkbase_refs[role_type] = []

                        # Resolve href relative to XSD URL
                        resolved_href = resolve_url(xsd_url, href)
                        linkbase_refs[role_type].append((resolved_href, role))

        return XsdSchema(
            target_namespace=target_namespace,
            prefix=schema_prefix,
            elements=elements,
            linkbase_refs=linkbase_refs,
            imports=imports,
        )

    def _parse_linkbase(
        self,
        linkbase_url: str,
        link_tag_suffix: str,
        arc_tag_suffix: str,
        additional_parser: Optional[XmlParser] = None,
    ) -> Iterator[Tuple[str, str, str, str, Dict[str, QName], ET.Element]]:
        """
        Generic method to parse linkbase files.

        Args:
            linkbase_url: URL of the linkbase file
            link_tag_suffix: Suffix for link tag (e.g., "presentationLink", "labelLink")
            arc_tag_suffix: Suffix for arc tag (e.g., "presentationArc", "labelArc")
            additional_parser: Optional additional parser for specific elements

        Yields:
            Tuple of (to_label, from_label, arc_arcrole, link_role, label_to_qname, elem)
        """
        # Fetch the linkbase file
        linkbase_path = self.file_loader.fetch(linkbase_url)
        if linkbase_path is None:
            return  # Skip if cannot fetch

        # Mapping from label to QName for this linkbase
        label_to_qname: Dict[str, QName] = {}

        link_role = None
        link_tag = None
        loc_tag = None
        arc_tag = None
        href_attr_name = None
        role_attr_name = None
        arcrole_attr_name = None
        label_attr_name = None
        to_attr_name = None
        from_attr_name = None
        type_attr_name = None

        for event, elem in ET.iterparse(
            str(linkbase_path), events=["start-ns", "start", "end"]
        ):
            # Let additional parser process events
            if additional_parser:
                additional_parser.on_xml_event(event, elem)

            if event == "start-ns":
                prefix, uri = elem
                if prefix == "link":
                    link_tag = f"{{{uri}}}{link_tag_suffix}"
                    loc_tag = f"{{{uri}}}loc"
                    arc_tag = f"{{{uri}}}{arc_tag_suffix}"
                elif prefix == "xlink":
                    href_attr_name = f"{{{uri}}}href"
                    role_attr_name = f"{{{uri}}}role"
                    arcrole_attr_name = f"{{{uri}}}arcrole"
                    label_attr_name = f"{{{uri}}}label"
                    to_attr_name = f"{{{uri}}}to"
                    from_attr_name = f"{{{uri}}}from"
                    type_attr_name = f"{{{uri}}}type"

            if link_tag and loc_tag and arc_tag:
                if event == "start":
                    if elem.tag == link_tag:
                        if link_role is not None:
                            raise ValueError(
                                f"Nested {link_tag_suffix} elements found."
                            )
                        link_role = elem.get(role_attr_name)
                elif event == "end":
                    if elem.tag == link_tag:
                        link_role = None
                        # Clear label mappings for next link
                        label_to_qname = {}
                    elif elem.tag == loc_tag:
                        # Check if this is a locator type
                        xlink_type = (
                            elem.get(type_attr_name) if type_attr_name else None
                        )

                        if xlink_type == "locator":
                            loc_role = (
                                elem.get(role_attr_name) if role_attr_name else None
                            )

                            # Locators that carry an xlink:role attribute point to
                            # remote resources (e.g. label definition files) rather
                            # than concepts defined in the current schema. These are
                            # handled separately by LabelResourceParser, so skip
                            # QName resolution here.
                            if loc_role:
                                continue

                            # For locator type, both href and label are required
                            loc_href = elem.get(href_attr_name)
                            loc_label = elem.get(label_attr_name)

                            if not loc_href:
                                raise ValueError(
                                    "link:loc with xlink:type='locator' is missing required xlink:href attribute"
                                )
                            if not loc_label:
                                raise ValueError(
                                    "link:loc with xlink:type='locator' is missing required xlink:label attribute"
                                )

                            # Parse the href to extract schema URL and element ID
                            if "#" in loc_href:
                                schema_url, element_id = loc_href.split("#", 1)

                                # Resolve relative URLs against the linkbase URL
                                resolved_schema_url = resolve_url(
                                    linkbase_url, schema_url
                                )

                                # Get or parse the XSD schema
                                schema = self._get_or_parse_xsd_schema(
                                    resolved_schema_url
                                )

                                # Look up the element in the schema
                                if schema and element_id in schema.elements:
                                    element_name = schema.elements[element_id]
                                    # Create QName
                                    qname = QName(
                                        local_name=element_name,
                                        namespace_uri=schema.target_namespace,
                                        prefix=schema.prefix,
                                    )
                                    label_to_qname[loc_label] = qname

                    elif elem.tag == arc_tag:
                        to_label = elem.get(to_attr_name)
                        from_label = elem.get(from_attr_name)
                        arc_arcrole = elem.get(arcrole_attr_name)

                        if to_label and from_label:
                            yield (
                                to_label,
                                from_label,
                                arc_arcrole or "",
                                link_role or "",
                                label_to_qname,
                                elem,
                            )

    def load_presentation_links(self) -> Iterator[Arc]:
        """
        Load and yield presentation arcs from presentation linkbase files.

        Note: This method may yield duplicate arcs if the same relationship
        appears multiple times in the linkbase files. It is the caller's
        responsibility to deduplicate arcs if needed.

        Yields:
            Arc objects with presentation relationships
        """
        self._ensure_linkbase_refs()
        if "presentationLinkbaseRef" not in self.linkbase_refs:
            return  # No presentation linkbase found

        for href, _ in self.linkbase_refs["presentationLinkbaseRef"]:
            # Parse linkbase file using the generic parser
            for (
                to_label,
                from_label,
                arc_arcrole,
                link_role,
                label_to_qname,
                _,
            ) in self._parse_linkbase(href, "presentationLink", "presentationArc"):
                # Resolve labels to QNames
                from_qname = label_to_qname.get(from_label)
                to_qname = label_to_qname.get(to_label)

                if from_qname and to_qname:
                    yield Arc(
                        from_qname=from_qname,
                        to_qname=to_qname,
                        role=link_role,
                        arcrole=arc_arcrole,
                        weight=None,
                    )

    def load_calculation_links(self) -> Iterator[Arc]:
        """
        Load and yield calculation arcs from calculation linkbase files.

        Note: This method may yield duplicate arcs if the same relationship
        appears multiple times in the linkbase files. It is the caller's
        responsibility to deduplicate arcs if needed.

        Yields:
            Arc objects with calculation relationships and weights
        """
        self._ensure_linkbase_refs()
        if "calculationLinkbaseRef" not in self.linkbase_refs:
            return  # No calculation linkbase found

        for href, _ in self.linkbase_refs["calculationLinkbaseRef"]:
            # Parse linkbase file using the generic parser
            for (
                to_label,
                from_label,
                arc_arcrole,
                link_role,
                label_to_qname,
                elem,
            ) in self._parse_linkbase(href, "calculationLink", "calculationArc"):
                # Resolve labels to QNames
                from_qname = label_to_qname.get(from_label)
                to_qname = label_to_qname.get(to_label)

                if from_qname and to_qname:
                    # Parse weight for calculation arcs
                    weight = None
                    weight_str = elem.get("weight")
                    if weight_str:
                        try:
                            weight = Decimal(weight_str)
                        except (ValueError, TypeError, InvalidOperation):
                            weight = None

                    yield Arc(
                        from_qname=from_qname,
                        to_qname=to_qname,
                        role=link_role,
                        arcrole=arc_arcrole,
                        weight=weight,
                    )

    def load_labels(self) -> Iterator[Label]:
        """
        Load and yield labels from label linkbase files.

        Returns labels in all languages (both "ja" and "en").
        Users should filter by language as needed.

        Note: This method may yield duplicate labels if the same label
        appears multiple times in the linkbase files. It is the caller's
        responsibility to deduplicate labels if needed.

        Yields:
            Label objects with text, language, role and arcrole information
        """
        self._ensure_linkbase_refs()
        if "labelLinkbaseRef" not in self.linkbase_refs:
            return  # No label linkbase found

        for href, _ in self.linkbase_refs["labelLinkbaseRef"]:
            # Create label resource parser with context for resolving remote labels
            label_parser = LabelResourceParser(
                file_loader=self.file_loader,
                base_url=href,  # Use original href URL for resolving remote labels
            )

            # Parse linkbase file using the generic parser with label parser
            for (
                to_label,
                from_label,
                arc_arcrole,
                link_role,
                label_to_qname,
                arc_elem,
            ) in self._parse_linkbase(href, "labelLink", "labelArc", label_parser):
                # Get QName for the concept
                qname = label_to_qname.get(from_label)
                # Get label resources from the label parser
                labels = label_parser.label_resources.get(to_label, [])

                if qname and labels:
                    label_priority = None
                    if arc_elem is not None:
                        priority_attr = arc_elem.get("priority")
                        if priority_attr is not None:
                            try:
                                label_priority = Decimal(priority_attr)
                            except (InvalidOperation, ValueError, TypeError):
                                label_priority = None
                    for label_text, label_lang, label_role in labels:
                        yield Label(
                            qname=qname,
                            text=label_text,
                            lang=label_lang,
                            link_role=link_role,
                            label_role=label_role,
                            arcrole=arc_arcrole,
                            priority=label_priority,
                        )

    def _collect_all_schemas_with_imports(
        self, root_schema_urls: List[str]
    ) -> Set[str]:
        """
        Recursively collect all XSD schemas including imports.

        Args:
            root_schema_urls: List of root XSD schema URLs to start from

        Returns:
            Set of all schema URLs including imported schemas
        """
        all_schemas = set()
        to_process = list(root_schema_urls)

        while to_process:
            schema_url = to_process.pop(0)

            # Skip if already processed (prevents cycles)
            if schema_url in all_schemas:
                continue

            all_schemas.add(schema_url)

            # Parse the schema and get its imports
            schema = self._get_or_parse_xsd_schema(schema_url)
            if schema and schema.imports:
                # Add imports to processing queue
                for import_url in schema.imports:
                    if import_url not in all_schemas:
                        to_process.append(import_url)

        return all_schemas

    def _ensure_linkbase_refs(self) -> None:
        """Ensure linkbase references are loaded and merged from all XSD schemas."""
        if not self.linkbase_refs:
            linkbase_refs = {}
            root_xsd_files = self.find_xsd_files()

            # Determine which schemas to process
            if self.follow_xsd_imports:
                # Collect all schemas including imports
                schemas_to_process = self._collect_all_schemas_with_imports(
                    [str(xsd_path) for xsd_path in root_xsd_files]
                )
            else:
                # Only process root schemas (backward compatible)
                schemas_to_process = [str(xsd_path) for xsd_path in root_xsd_files]

            # Merge linkbase refs from all schemas
            for schema_url in schemas_to_process:
                # Get or parse the schema
                schema = self._get_or_parse_xsd_schema(schema_url)
                if schema and schema.linkbase_refs:
                    for role_type, refs_list in schema.linkbase_refs.items():
                        if role_type not in linkbase_refs:
                            linkbase_refs[role_type] = []
                        linkbase_refs[role_type].extend(refs_list)

            self.linkbase_refs = linkbase_refs