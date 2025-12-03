import * as React from 'react';
import { ILepineSearchResult } from '../models/ISearchResult';
// Fix TS2305: Move ImageFit to Image import
import { ImageFit } from '@fluentui/react/lib/Image'; 
import { DocumentCard, DocumentCardPreview, DocumentCardTitle, DocumentCardActivity, IDocumentCardPreviewProps } from '@fluentui/react/lib/DocumentCard';
import { Spinner, SpinnerSize, Text } from '@fluentui/react';

export interface IContentProps {
  items: ILepineSearchResult[];
  isLoading: boolean;
}

export default class LepineSearchResultsContent extends React.Component<IContentProps, {}> {
  public render() {
    const { items, isLoading } = this.props;

    if (isLoading) {
      return <Spinner size={SpinnerSize.large} label="Loading documents..." />;
    }

    return (
      <div>
        <Text variant="large" block style={{marginBottom: 15}}>
          Found {items.length} results
        </Text>
        
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '20px' }}>
          {items.map(item => {
            const previewProps: IDocumentCardPreviewProps = {
              previewImages: [
                {
                  name: item.name,
                  previewImageSrc: item.thumbnailUrl,
                  imageFit: ImageFit.cover,
                  width: 318,
                  height: 196,
                  iconSrc: '',
                },
              ],
            };

            return (
              <DocumentCard 
                key={item.id} 
                onClickHref={item.href}
                onClickTarget="_blank"
                styles={{ root: { maxWidth: 318, minWidth: 318 } }}
              >
                <DocumentCardPreview {...previewProps} />
                <DocumentCardTitle title={item.name} shouldTruncate={false} />
                <DocumentCardActivity 
                  activity="Created" 
                  people={[{ name: item.location, profileImageSrc: '' }]} 
                />
              </DocumentCard>
            );
          })}
        </div>
      </div>
    );
  }
}