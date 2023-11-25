import * as React from 'react';
import { ISummariseProps } from './ISummariseProps';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { FontIcon } from '@fluentui/react/lib/Icon';
import { summariseStyles, loadingSpinnerStyles } from './styles';
import { isEmpty } from '@microsoft/sp-lodash-subset';
import { useAzureFunctions, useSharePointRest, useMicrosoftGraph } from '../../../hooks';
import { SUMMARY_COLUMN_NAME, SITE_PAGES_LIBRARY_NAME } from '../../../constants/constants';
import { Sentence } from './types';

const Summarise: React.FC<ISummariseProps> = (props) => {

  const [loading, setLoading] = React.useState<boolean>(true);
  const [sentences, setSentences] = React.useState<Sentence[]>([]);

  const { aadHttpClientFactory, msGraphClientFactory, spHttpClient, siteId, pageId, siteUrl, pageItemId } = props;
  const { getSummaryUsingOpenAI, updatePagePnPPowerShell } = useAzureFunctions(aadHttpClientFactory);
  const { callMicrosoftGraphAPI } = useMicrosoftGraph(msGraphClientFactory);
  const { getItem } = useSharePointRest(spHttpClient, siteUrl);

  const cleanPageContent = (pageContent: string): string => {

    //remove html tags from the content
    pageContent = pageContent.replace(/<[^>]*>?/gm, '');

    //replace " with '
    pageContent = pageContent.replace(/"/g, "'");

    // remove all unicode characters
    pageContent = pageContent.replace(/[^\x00-\x7F]/g, "");

    return pageContent;
  };

  const getPageContentUsingGraphAPI = async (): Promise<string> => {

    // get the page content from the Microsoft Graph API
    const response = await callMicrosoftGraphAPI(
      "get",
      `/sites/${siteId}/pages/${pageId}/microsoft.graph.sitepage/webparts`,
      "beta",
      null,
      null,
      null,
      "(isof('microsoft.graph.textWebPart'))",
      null
    );
    return response.value?.map((webPart: any) => webPart.innerHtml)?.join(' ') || '';
  };

  const getSummaryFromPage = async (): Promise<Sentence[]> => {

    // get the summary from the page
    const page = await getItem(SITE_PAGES_LIBRARY_NAME, pageItemId, [SUMMARY_COLUMN_NAME]);

    // if page is empty, return
    if (page === undefined) {
      return [];
    }

    let summary: Sentence[] = null;

    // if summary is not empty, return the summary by splitting it into sentences
    if (!isEmpty(page) && !isEmpty(page[SUMMARY_COLUMN_NAME])) {

      /* const summaryFromColumn = page[SUMMARY_COLUMN_NAME];
      const parsedSummary = JSON.parse(summaryFromColumn);
      summary = Object.entries(parsedSummary).map(([language, content]: [string, string]): Sentence => ({
        language,
        content
      })); */

      summary = JSON.parse(page[SUMMARY_COLUMN_NAME]);
    }

    // wait for 1 second before returning the summary to show the loading spinner
    await new Promise((resolve) => setTimeout(resolve, 1000));

    return summary;

  };

  const getSummaryFromAPI = async (): Promise<Sentence[]> => {

    let pageContent = await getPageContentUsingGraphAPI();

    // if page content is empty, return
    if (isEmpty(pageContent)) {
      return [];
    }

    // clean the page content
    pageContent = cleanPageContent(pageContent);

    // get summary from OpenAI
    const summary = await getSummaryUsingOpenAI(pageContent);

    // if summary is empty, return
    if (isEmpty(summary)) {
      return [];
    }

    return summary && Object.keys(summary).map((language: string) => ({ language, content: summary[language] }));
  };

  const executeSummaryTasksAndUpdatePage = async (): Promise<void> => {
    let summary: Sentence[] = await getSummaryFromPage();
    if (summary === null) {
      summary = await getSummaryFromAPI();

      if (!isEmpty(summary)) {
        // update the page with the summary
        updatePagePnPPowerShell(siteUrl, pageItemId, SUMMARY_COLUMN_NAME, JSON.stringify(summary));
      }
    }
    // covertSummaryToSentences(summary);
    setSentences(summary);
  };

  React.useEffect(() => {
    executeSummaryTasksAndUpdatePage()
      .then(
        () => setLoading(false)
      )
      .catch(
        (error) => {
          console.log("error", error);
          setSentences([]);
          setLoading(false);
        }
      );
  }, []);

  return (

    <div className={summariseStyles.mainContainer}>
      <div className={summariseStyles.titleContainer}>
        <FontIcon className={summariseStyles.icon} iconName="AlignLeft" />
        <span className={summariseStyles.title}>Summary</span>
      </div>
      {loading ? (
        <Spinner size={SpinnerSize.large} label="Loading summary..." styles={loadingSpinnerStyles} />
      ) : isEmpty(sentences) ? (
        <p className={summariseStyles.description}>No summary available</p>
      ) : (
        sentences.map((sentence, index) => (
          <p className={summariseStyles.descriptionContainer} key={index}>
            <span className={`${summariseStyles.description}`}><span className={summariseStyles.language}>{sentence.language}: </span>{sentence.content}</span>
          </p>
        ))
      )}
    </div>

  );
}

export default Summarise;