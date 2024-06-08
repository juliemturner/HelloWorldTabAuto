import { useEffect, useState } from "react";
import { IDocumentsByType, IFile } from "../common/models";

const useDocuments = () => {
  const LOG_SOURCE = "useDocuments";

  const [documents, setDocuments] = useState<IFile[]>([]);
  const [documentsByType, setDocumentsByType ] = useState<IDocumentsByType>({});
  const [totalSize, setTotalSize] = useState<number>(0);
  const [isError, setError] = useState<Boolean>(false);

  //side effect with empty dependency array, means to run once
  useEffect(() => {
    (async () => {
      try {
        const itemsResult = await fetch("https://afteamssamples.azurewebsites.net/api/getDocuments");
        if(itemsResult.ok){
          console.log(`${LOG_SOURCE} Retrieved files`);
          const items: IFile[] = await itemsResult.json();
          console.log(`${LOG_SOURCE} Resolved JSON: ${items.length} items`);
          setDocuments(items);          
          console.log(`${LOG_SOURCE} Set items`);
          setTotalSize(
            items.length > 0
              ? items.reduce<number>((acc: number, item: IFile) => {
                  return acc + Number(item.Size);
                }, 0)
              : 0
          );
          console.log(`${LOG_SOURCE} Set total size`);
          const dbt: IDocumentsByType = {};
          items.forEach((o) => {
            const ext = o.Name.substring(o.Name?.lastIndexOf("."));
            if (dbt[ext] == null){
              dbt[ext] = [o];
            }else{
              dbt[ext].push(o);
            }
          });
          setDocumentsByType(dbt);
          console.log(`${LOG_SOURCE} Set documents by type`);
        }else{
          console.error(`${LOG_SOURCE} could not fetch documents ${itemsResult.status}:${itemsResult.statusText}`);
        }
      } catch (err) {
        setError(true);
        console.error(`${LOG_SOURCE} (getting files useEffect) - ${JSON.stringify(err)}`);
      }
    })();
  }, []);

  // const updateDocuments = async () => {
  //   try {
  //     const [batchedSP, execute] = _sp.batched();

  //     //clone documents
  //     const items = JSON.parse(JSON.stringify(documents));

  //     const res: IItemUpdateResult[] = [];

  //     for (let i = 0; i < items.length; i++) {
  //       // you need to use .then syntax here as otherwise the application will stop and await the result
  //       batchedSP.web.lists
  //         .getByTitle(LIBRARY_NAME)
  //         .items.getById(items[i].Id)
  //         .update({ Title: `${items[i].Name}-Updated` })
  //         .then((r) => res.push(r));
  //     }
  //     // Executes the batched calls
  //     await execute();

  //     // Results for all batched calls are available
  //     for (let i = 0; i < res.length; i++) {
  //       //If the result is successful update the item
  //       //NOTE: This code is over simplified, you need to make sure the Id's match
  //       const item = await res[i].item.select("Id, Title")<{
  //         Id: number;
  //         Title: string;
  //       }>();
  //       items[i].Name = item.Title;
  //     }

  //     //Update the state
  //     setDocuments(items);
  //     setError(false);
  //   } catch (err) {
  //     setError(true);
  //     Logger.write(
  //       `${LOG_SOURCE} (updating titles) - ${JSON.stringify(err)} - `,
  //       LogLevel.Error
  //     );
  //   }
  // };

  return {documents, documentsByType, totalSize, isError} as const;
};

export default useDocuments;