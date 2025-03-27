import * as React from "react";
import { SPFI, spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { DetailsList, IColumn, SelectionMode } from "@fluentui/react";
import styles from "./BillFolderSearch.module.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBillFolderSearchProps {
  context: WebPartContext; // SPFx context passed from the web part
}

interface ISharePointItem {
  name: string;
  url: string;
  type: "Folder" | "File";
  created?: string;
  modified?: string;
}

const Pagination: React.FC<{
  totalItems: number;
  itemsPerPage: number;
  currentPage: number;
  onPageChange: (page: number) => void;
}> = ({ totalItems, itemsPerPage, currentPage, onPageChange }) => {
  const totalPages = Math.ceil(totalItems / itemsPerPage);

  const handleClick = (page: number) => {
    if (page >= 1 && page <= totalPages) {
      onPageChange(page);
    }
  };

  return (
    <div className={styles.pagination}>
      <button disabled={currentPage === 1} onClick={() => handleClick(currentPage - 1)}>
        Previous
      </button>
      {Array.from({ length: totalPages }, (_, index) => (
        <button
          key={index + 1}
          onClick={() => handleClick(index + 1)}
          className={currentPage === index + 1 ? styles.activePage : ""}
        >
          {index + 1}
        </button>
      ))}
      <button disabled={currentPage === totalPages} onClick={() => handleClick(currentPage + 1)}>
        Next
      </button>
    </div>
  );
};

const BillFolderSearch: React.FC<IBillFolderSearchProps> = (props) => {
  const [searchTerm, setSearchTerm] = React.useState<string>("");
  const [results, setResults] = React.useState<ISharePointItem[]>([]);
  const [filteredResults, setFilteredResults] = React.useState<ISharePointItem[]>([]);
  const [error, setError] = React.useState<string>("");
  const [currentPage, setCurrentPage] = React.useState<number>(1);
  const itemsPerPage = 10;

  const sp: SPFI = spfi().using(SPFx(props.context));

  const fetchAllItems = async (): Promise<void> => {
    setError("");

    try {
      const items = await sp.web.lists
        .getByTitle("legalbill")
        .items.select("Id", "Title", "FileRef", "FileSystemObjectType", "Created", "Modified", "FileLeafRef")
        .top(5000)();

      const mappedResults = items.map((item) => ({
        name: item.FileLeafRef || item.Title || `Item ${item.Id}`,
        url: item.FileRef,
        type: item.FileSystemObjectType === 1 ? ("Folder" as const) : ("File" as const), // Explicit type assertion
        created: item.Created,
        modified: item.Modified,
      }));

      setResults(mappedResults);
      setFilteredResults(mappedResults.slice(0, itemsPerPage)); // Set initial page results
    } catch (err) {
      console.error("Error fetching items:", err);
      setError("Failed to fetch items. Please check permissions or configuration.");
    }
  };

  const searchItems = (): void => {
    const searchLower = searchTerm.toLowerCase(); // Normalize search term to lowercase
    const filtered = results.filter(
      (item) =>
        (item.name && item.name.toLowerCase().includes(searchLower)) ||
        (item.url && item.url.toLowerCase().includes(searchLower))
    );

    // Update filtered results and reset to the first page
    setFilteredResults(filtered.slice(0, itemsPerPage));
    setCurrentPage(1);

    // Debugging: Log search results
    console.log("Search results:", filtered);
  };

  const handlePageChange = (page: number): void => {
    setCurrentPage(page);
    const startIndex = (page - 1) * itemsPerPage;
    const paginatedResults = filteredResults.slice(startIndex, startIndex + itemsPerPage);
    setFilteredResults(paginatedResults);
  };

  React.useEffect(() => {
    fetchAllItems(); // Fetch items when the component mounts
  }, []);

  const columns: IColumn[] = [
    { key: "name", name: "Name", fieldName: "name", minWidth: 200, isResizable: true },
    { key: "type", name: "Type", fieldName: "type", minWidth: 100, isResizable: true },
    {
      key: "created",
      name: "Created",
      fieldName: "created",
      minWidth: 150,
      isResizable: true,
      onRender: (item: ISharePointItem) =>
        item.created ? new Date(item.created).toLocaleDateString() : "",
    },
    {
      key: "modified",
      name: "Modified",
      fieldName: "modified",
      minWidth: 150,
      isResizable: true,
      onRender: (item: ISharePointItem) =>
        item.modified ? new Date(item.modified).toLocaleDateString() : "",
    },
    {
      key: "url",
      name: "URL",
      fieldName: "url",
      minWidth: 300,
      isResizable: true,
      onRender: (item: ISharePointItem) => (
        <a href={item.url} target="_blank" rel="noopener noreferrer">
          Open
        </a>
      ),
    },
  ];

  return (
    <div className={styles.billFolderSearch}>
      <h2>Legal Bills Library</h2>
      <input
        type="text"
        placeholder="Search"
        value={searchTerm}
        onChange={(e) => setSearchTerm(e.target.value)}
        className={styles.searchBox}
      />
      <button onClick={searchItems} className={styles.searchButton}>
        Search
      </button>
      {error && <p className={styles.error}>{error}</p>}
      <DetailsList
        items={filteredResults}
        columns={columns}
        selectionMode={SelectionMode.none}
        isHeaderVisible={true}
      />
      <Pagination
        totalItems={filteredResults.length}
        itemsPerPage={itemsPerPage}
        currentPage={currentPage}
        onPageChange={handlePageChange}
      />
    </div>
  );
};

export default BillFolderSearch;
