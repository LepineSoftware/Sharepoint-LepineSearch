import * as React from 'react';
import { ILepineSearchResult } from '../models/ISearchResult';
import { ImageFit } from '@fluentui/react/lib/Image'; 
import { DetailsList, DetailsListLayoutMode, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { DocumentCard, DocumentCardTitle } from '@fluentui/react/lib/DocumentCard';
import { 
  Spinner, 
  SpinnerSize, 
  Link, 
  Icon, 
  Stack, 
  DefaultButton, 
  PrimaryButton, 
  Text,
  Modal,
  IconButton,
  Image,
  IIconProps
} from '@fluentui/react';

export interface IContentProps {
  items: ILepineSearchResult[];
  isLoading: boolean;
  isCardView: boolean;
}

interface IContentState {
  currentPage: number;
  selectedItem: ILepineSearchResult | null;
}

const cancelIcon: IIconProps = { iconName: 'Cancel' };
const openIcon: IIconProps = { iconName: 'OpenInNewWindow' };
const prevIcon: IIconProps = { iconName: 'ChevronLeft' };
const nextIcon: IIconProps = { iconName: 'ChevronRight' };

export default class LepineSearchResultsContent extends React.Component<IContentProps, IContentState> {

  constructor(props: IContentProps) {
    super(props);
    this.state = {
      currentPage: 1,
      selectedItem: null
    };
  }

  public componentDidUpdate(prevProps: IContentProps) {
    if (prevProps.items !== this.props.items || prevProps.isCardView !== this.props.isCardView) {
      this.setState({ currentPage: 1 });
    }
  }

  // --- NAVIGATION LOGIC ---

  private _onItemClick = (item: ILepineSearchResult) => {
    this.setState({ selectedItem: item });
  }

  private _closeModal = () => {
    this.setState({ selectedItem: null });
  }

  private _onNavigate = (direction: 'next' | 'prev') => {
    const { items } = this.props;
    const { selectedItem } = this.state;
    if (!selectedItem) return;

    const currentIndex = items.indexOf(selectedItem);
    let newIndex = direction === 'next' ? currentIndex + 1 : currentIndex - 1;

    if (newIndex >= 0 && newIndex < items.length) {
      this.setState({ selectedItem: items[newIndex] });
    }
  }

  private _onKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'ArrowRight') this._onNavigate('next');
    if (e.key === 'ArrowLeft') this._onNavigate('prev');
  }

  // --- HELPERS (FIXED) ---

  private _getHighResPreviewUrl(item: ILepineSearchResult): string {
    const url = item.thumbnailUrl;

    // Case 1: Modern Drive API - Upgrade 'large' to HD (1080p)
    if (url.indexOf('/large/content') > -1) {
        return url.replace('/large/content', '/c1920x1080/content');
    }

    // Case 2: Legacy getpreview.ashx - Upgrade resolution
    if (url.indexOf('resolution=') > -1) {
        return url.replace('resolution=0', 'resolution=3');
    }

    // Fallback: Append resolution if missing
    if (url.indexOf('getpreview.ashx') > -1 && url.indexOf('resolution=') === -1) {
        return url + (url.indexOf('?') > -1 ? '&' : '?') + "resolution=3";
    }

    return url;
  }

  private _columns: IColumn[] = [
    {
      key: 'icon', name: '', fieldName: 'fileType', minWidth: 20, maxWidth: 20,
      onRender: () => <Icon iconName="Page" />
    },
    {
      key: 'name', name: 'Name', fieldName: 'name', minWidth: 150, maxWidth: 300, isResizable: true,
      onRender: (item: ILepineSearchResult) => (
        <Link onClick={() => this._onItemClick(item)}>{item.name}</Link>
      )
    },
    {
      key: 'tags', name: 'Tags', fieldName: 'tags', minWidth: 100, maxWidth: 200,
      onRender: (item: ILepineSearchResult) => <span>{item.tags.join(', ')}</span>
    },
    {
      key: 'fileType', name: 'File Type', fieldName: 'fileType', minWidth: 80, maxWidth: 100, isResizable: true,
      onRender: (item: ILepineSearchResult) => <span>{item.fileType ? item.fileType.toUpperCase() : ''}</span>
    }
  ];

  public render() {
    const { items, isLoading, isCardView } = this.props;
    const { currentPage, selectedItem } = this.state;

    // --- TYPE CHECKERS ---
    const isVideoFile = (fileType: string) => 
        ['mp4', 'mov', 'webm', 'ogg', 'mkv', 'avi'].indexOf((fileType || '').toLowerCase()) > -1;

    // --- MODAL STATE HELPERS ---
    const isSelectedItemVideo = selectedItem ? isVideoFile(selectedItem.fileType) : false;
    
    const selectedIndex = selectedItem ? items.indexOf(selectedItem) : -1;
    const hasPrev = selectedIndex > 0;
    const hasNext = selectedIndex < items.length - 1;

    if (isLoading) {
      return <Spinner size={SpinnerSize.large} label="Loading documents..." />;
    }

    const pageSize = isCardView ? 12 : 25;
    const totalPages = Math.ceil(items.length / pageSize);
    const startIndex = (currentPage - 1) * pageSize;
    const currentItems = items.slice(startIndex, startIndex + pageSize);

    const previewContainerStyle: React.CSSProperties = {
        position: 'relative',
        width: '100%',
        height: '130px',
        backgroundColor: '#f3f2f1',
        overflow: 'hidden',
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center'
    };

    return (
      <Stack tokens={{ childrenGap: 20 }}>
        
        {/* VIEW AREA */}
        <div>
          {!isCardView ? (
            <DetailsList
              items={currentItems}
              columns={this._columns}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              compact={true}
            />
          ) : (
            <div style={{ 
              display: 'grid', 
              gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', 
              gap: '20px',
              alignItems: 'stretch' 
            }}>
              {currentItems.map(item => {
                const isVideo = isVideoFile(item.fileType);

                return (
                  <DocumentCard 
                    key={item.id} 
                    onClick={() => this._onItemClick(item)}
                    styles={{ root: { width: '100%', height: '100%', minWidth: 'auto', cursor: 'pointer' } }}
                  >
                    {/* UNIFIED PREVIEW CONTAINER */}
                    <div style={previewContainerStyle}>
                        {isVideo ? (
                            // VIDEO CARD RENDER
                            <>
                            {/* Use Image with thumbnail poster */}
                            <Image
                                src={item.thumbnailUrl} 
                                imageFit={ImageFit.cover}
                                width="100%"
                                height={130}
                                alt={`Preview of ${item.name}`}
                            />
                            <div style={{ position: 'absolute', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', color: 'rgba(255,255,255,0.9)', pointerEvents: 'none', zIndex: 2 }}>
                                <Icon iconName="Play" styles={{ root: { fontSize: 32, filter: 'drop-shadow(0 0 4px rgba(0,0,0,0.5))' } }} />
                            </div>
                            </>
                        ) : (
                            // STANDARD IMAGE RENDER
                            <Image
                                src={item.thumbnailUrl}
                                imageFit={ImageFit.centerCover} 
                                width="100%"
                                height={130}
                                alt={`Preview of ${item.name}`}
                                onError={(ev) => {
                                    // Fallback: hide broken image, effectively showing grey background
                                    (ev.target as HTMLImageElement).style.visibility = 'hidden';
                                }}
                            />
                        )}
                    </div>

                    <DocumentCardTitle 
                        title={item.name} 
                        shouldTruncate={true}
                        styles={{ root: { padding: '8px 12px', height: '38px', fontSize: '12px', lineHeight: '16px', fontWeight: 'normal', overflow: 'hidden' } }}
                    />
                  </DocumentCard>
                );
              })}
            </div>
          )}
        </div>

        {/* PAGINATION */}
        {items.length > pageSize && (
          <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 20 }} styles={{ root: { marginTop: 20 } }}>
            <DefaultButton 
              text="Previous" 
              onClick={() => this.setState({ currentPage: currentPage - 1 })}
              disabled={currentPage === 1}
              iconProps={{ iconName: 'ChevronLeft' }}
            />
            <Text variant="mediumPlus" styles={{ root: { alignSelf: 'center' } }}>
              Page {currentPage} of {totalPages}
            </Text>
            <DefaultButton 
              text="Next" 
              onClick={() => this.setState({ currentPage: currentPage + 1 })}
              disabled={currentPage === totalPages}
              menuIconProps={{ iconName: 'ChevronRight' }} 
            />
          </Stack>
        )}

        {/* ITEM PREVIEW MODAL */}
        <Modal
          isOpen={!!selectedItem}
          onDismiss={this._closeModal}
          isBlocking={false}
          containerClassName="lepineSearchModalContainer"
        >
          {selectedItem && (
            <div 
                style={{ padding: '20px', maxWidth: '850px', width: '90vw', outline: 'none' }}
                onKeyDown={this._onKeyDown}
                tabIndex={0}
            >
              
              <Stack horizontal horizontalAlign="space-between" verticalAlign="start" styles={{ root: { marginBottom: 15 } }}>
                <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>{selectedItem.name}</Text>
                <IconButton iconProps={cancelIcon} onClick={this._closeModal} ariaLabel="Close" />
              </Stack>

              {/* MEDIA & NAVIGATION CONTAINER */}
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} styles={{ root: { marginBottom: 20 } }}>
                 
                 <IconButton 
                    iconProps={prevIcon} 
                    disabled={!hasPrev} 
                    onClick={() => this._onNavigate('prev')}
                    styles={{ icon: { fontSize: 24, fontWeight: 'bold' } }}
                    ariaLabel="Previous Item"
                 />

                 <div style={{ 
                      backgroundColor: '#f3f2f1', 
                      display: 'flex', 
                      justifyContent: 'center', 
                      alignItems: 'center',
                      minHeight: '300px',
                      flexGrow: 1, 
                      position: 'relative',
                      borderRadius: '4px',
                      overflow: 'hidden'
                  }}>
                     {isSelectedItemVideo ? (
                        <video 
                            controls 
                            autoPlay 
                            poster={this._getHighResPreviewUrl(selectedItem)}
                            src={selectedItem.href}
                            style={{ maxWidth: '100%', maxHeight: '500px', outline: 'none' }}
                        />
                     ) : (
                        <Image 
                            src={this._getHighResPreviewUrl(selectedItem)} 
                            alt={`Preview of ${selectedItem.name}`}
                            imageFit={ImageFit.contain}
                            width={600}
                            height={400}
                            styles={{ root: { maxHeight: '500px' } }}
                        />
                     )}
                  </div>

                  <IconButton 
                    iconProps={nextIcon} 
                    disabled={!hasNext} 
                    onClick={() => this._onNavigate('next')}
                    styles={{ icon: { fontSize: 24, fontWeight: 'bold' } }}
                    ariaLabel="Next Item"
                 />
              </Stack>

              <Stack tokens={{ childrenGap: 5 }} styles={{ root: { marginBottom: 20 } }}>
                  <Text><strong>Type:</strong> {selectedItem.fileType ? selectedItem.fileType.toUpperCase() : 'Unknown'}</Text>
                  
                  <div style={{ height: '100px', overflowY: 'auto' }}>
                    <Text>
                        <strong>Tags:</strong> {selectedItem.tags && selectedItem.tags.length > 0 ? selectedItem.tags.join(', ') : 'None'}
                    </Text>
                  </div>

                  <Text variant="small">Item {selectedIndex + 1} of {items.length}</Text>
              </Stack>

              <Stack horizontal tokens={{ childrenGap: 10 }} horizontalAlign="end">
                  <DefaultButton text="Close" onClick={this._closeModal} />
                  <PrimaryButton 
                    text="Open in SharePoint" 
                    iconProps={openIcon}
                    href={selectedItem.href}
                    target="_blank"
                  />
              </Stack>
            </div>
          )}
        </Modal>

      </Stack>
    );
  }
}