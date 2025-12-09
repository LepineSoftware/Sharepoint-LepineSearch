import * as React from 'react';
import { ILepineSearchResult } from '../models/ISearchResult';
import { ImageFit } from '@fluentui/react/lib/Image'; 
import { DetailsList, DetailsListLayoutMode, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { DocumentCard } from '@fluentui/react/lib/DocumentCard';
import { 
  Link, 
  Icon, 
  Stack, 
  DefaultButton, 
  PrimaryButton, 
  Text, 
  Modal, 
  IconButton, 
  Image, 
  IIconProps, 
  Shimmer, 
  ShimmerElementType, 
  ShimmerElementsGroup, 
  Separator, 
  Label      
} from '@fluentui/react';
import { HighlightText } from './HighlightText';

const formatBytes = (bytes?: number) => {
    if (!bytes || isNaN(bytes) || bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
};

const formatDate = (dateStr?: string) => {
    if (!dateStr) return 'Unknown';
    return new Date(dateStr).toLocaleDateString(undefined, {
        year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit'
    });
};

export interface IContentProps {
  items: ILepineSearchResult[];
  isLoading: boolean;
  isCardView: boolean;
  searchQuery: string;
}

interface IContentState {
  currentPage: number;
  selectedItem: ILepineSearchResult | null;
}

const cancelIcon: IIconProps = { iconName: 'Cancel' };
const openIcon: IIconProps = { iconName: 'OpenInNewWindow' };
const downloadIcon: IIconProps = { iconName: 'Download' }; 
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
    if (prevProps.items !== this.props.items || 
        prevProps.isCardView !== this.props.isCardView || 
        prevProps.searchQuery !== this.props.searchQuery) {
      this.setState({ currentPage: 1 });
    }
  }

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

  private _getHighResPreviewUrl(item: ILepineSearchResult): string {
    const url = item.thumbnailUrl;
    if (url.indexOf('/large/content') > -1) {
        return url.replace('/large/content', '/c1920x1080/content');
    }
    if (url.indexOf('resolution=') > -1) {
        return url.replace('resolution=0', 'resolution=3');
    }
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
        <Link onClick={() => this._onItemClick(item)}>
           <HighlightText text={item.name} query={this.props.searchQuery} />
        </Link>
      )
    },
    {
      key: 'size', name: 'Size', fieldName: 'fileSize', minWidth: 70, maxWidth: 90, 
      onRender: (item: ILepineSearchResult) => <span>{formatBytes(item.fileSize)}</span>
    },
    {
      key: 'modified', name: 'Modified', fieldName: 'modified', minWidth: 120, maxWidth: 150, 
      onRender: (item: ILepineSearchResult) => <span>{new Date(item.modified || '').toLocaleDateString()}</span>
    },
    {
      key: 'tags', name: 'Tags', fieldName: 'tags', minWidth: 100, maxWidth: 200,
      onRender: (item: ILepineSearchResult) => <span>{item.tags.join(', ')}</span>
    }
  ];

  public render() {
    const { items, isLoading, isCardView, searchQuery } = this.props;
    const { currentPage, selectedItem } = this.state;

    // SHIMMER LOADING
    if (isLoading) {
      if (isCardView) {
        return (
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: '20px' }}>
             {[1,2,3,4,5,6,7,8].map(i => (
               <DocumentCard key={i} styles={{ root: { width: '100%', height: '180px' } }}>
                  <div style={{ padding: 10 }}>
                     <Shimmer customElementsGroup={
                       <div style={{ display: 'flex', flexDirection: 'column' }}>
                         <ShimmerElementsGroup shimmerElements={[{ type: ShimmerElementType.line, height: 120, width: '100%' }]} />
                         <div style={{ height: 10 }} />
                         <ShimmerElementsGroup shimmerElements={[{ type: ShimmerElementType.line, height: 16, width: '80%' }]} />
                       </div>
                     } />
                  </div>
               </DocumentCard>
             ))}
          </div>
        );
      } else {
        return (
          <Stack tokens={{ childrenGap: 20 }}>
             {[1,2,3,4,5,6,7,8,9,10].map(i => (
                <Shimmer key={i} />
             ))}
          </Stack>
        );
      }
    }

    const isVideoFile = (fileType: string) => 
        ['mp4', 'mov', 'webm', 'ogg', 'mkv', 'avi'].indexOf((fileType || '').toLowerCase()) > -1;

    const isSelectedItemVideo = selectedItem ? isVideoFile(selectedItem.fileType) : false;
    
    const selectedIndex = selectedItem ? items.indexOf(selectedItem) : -1;
    const hasPrev = selectedIndex > 0;
    const hasNext = selectedIndex < items.length - 1;

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
                    <div style={previewContainerStyle}>
                        {isVideo ? (
                            <>
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
                            <Image
                                src={item.thumbnailUrl}
                                imageFit={ImageFit.centerCover} 
                                width="100%"
                                height={130}
                                alt={`Preview of ${item.name}`}
                                onError={(ev) => {
                                    (ev.target as HTMLImageElement).style.visibility = 'hidden';
                                }}
                            />
                        )}
                    </div>

                    <div style={{ padding: '8px 12px', height: '38px', fontSize: '12px', lineHeight: '16px', overflow: 'hidden' }}>
                       <HighlightText text={item.name} query={searchQuery} />
                    </div>
                  </DocumentCard>
                );
              })}
            </div>
          )}
        </div>

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

        {/* --- ENHANCED PREVIEW MODAL --- */}
        <Modal
          isOpen={!!selectedItem}
          onDismiss={this._closeModal}
          isBlocking={false}
          containerClassName="lepineSearchModalContainer"
        >
          {selectedItem && (
            <div 
                style={{ padding: '0', width: '1000px', height: '80vh', display: 'flex', flexDirection: 'column', outline: 'none' }}
                onKeyDown={this._onKeyDown}
                tabIndex={0}
            >
              {/* Header */}
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center" styles={{ root: { padding: '15px 20px', borderBottom: '1px solid #eee' } }}>
                <Text variant="xLarge" styles={{ root: { fontWeight: 600, overflow:'hidden', whiteSpace:'nowrap', textOverflow:'ellipsis' } }} title={selectedItem.name}>
                    {selectedItem.name}
                </Text>
                <IconButton iconProps={cancelIcon} onClick={this._closeModal} ariaLabel="Close" />
              </Stack>

              {/* Body Split View */}
              <Stack horizontal styles={{ root: { flexGrow: 1, overflow: 'hidden' } }}>
                
                {/* Left: Preview Area */}
                <Stack 
                    verticalAlign="center" 
                    horizontalAlign="center" 
                    styles={{ root: { width: '680px', backgroundColor: '#1b1a19', position: 'relative', overflow:'hidden' } }}
                >
                    {/* Navigation Overlays */}
                    <div style={{ position: 'absolute', left: 10, zIndex: 10 }}>
                        <IconButton iconProps={prevIcon} disabled={!hasPrev} onClick={() => this._onNavigate('prev')} styles={{ root: { backgroundColor: 'rgba(255,255,255,0.8)', borderRadius: '50%' } }} />
                    </div>
                    <div style={{ position: 'absolute', right: 10, zIndex: 10 }}>
                         <IconButton iconProps={nextIcon} disabled={!hasNext} onClick={() => this._onNavigate('next')} styles={{ root: { backgroundColor: 'rgba(255,255,255,0.8)', borderRadius: '50%' } }} />
                    </div>

                    {isSelectedItemVideo ? (
                        <video 
                            controls 
                            autoPlay 
                            poster={this._getHighResPreviewUrl(selectedItem)}
                            src={selectedItem.href}
                            style={{ width: '100%', height: '100%', objectFit: 'contain', outline: 'none' }}
                        />
                     ) : (
                        <Image 
                            src={this._getHighResPreviewUrl(selectedItem)} 
                            alt={`Preview of ${selectedItem.name}`}
                            imageFit={ImageFit.contain}
                            styles={{ root: { width: '100%', height: '100%' } }}
                        />
                     )}
                </Stack>

                {/* Right: Details Panel */}
                <Stack styles={{ root: { width: '320px', padding: '20px', borderLeft: '1px solid #eee', overflowY: 'auto', backgroundColor: '#fff' } }} tokens={{ childrenGap: 15 }}>
                    
                    <Text variant="large" styles={{ root: { fontWeight: 600, marginBottom: 10 } }}>File Details</Text>
                    
                    <div>
                        <Label>Type</Label>
                        <Text>{selectedItem.fileType ? selectedItem.fileType.toUpperCase() : 'Unknown'}</Text>
                    </div>

                    <div>
                        <Label>Size</Label>
                        <Text>{formatBytes(selectedItem.fileSize)}</Text>
                    </div>

                    <div>
                        <Label>Modified</Label>
                        <Text>{formatDate(selectedItem.modified)}</Text>
                    </div>

                    <div>
                        <Label>Modified By</Label>
                        <Text>{selectedItem.modifiedBy || 'Unknown'}</Text>
                    </div>

                    {/* MOVED BUTTONS HERE */}
                    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 10 } }}>
                        <PrimaryButton 
                            text="Open in SharePoint" 
                            iconProps={openIcon}
                            href={selectedItem.href}
                            target="_blank"
                            styles={{ root: { width: '100%' } }}
                        />
                        <DefaultButton 
                            text="Download" 
                            iconProps={downloadIcon}
                            href={selectedItem.href}
                            download={selectedItem.name}
                            target="_blank"
                            styles={{ root: { width: '100%' } }}
                        />
                    </Stack>

                    <Separator />

                    <div>
                        <Label>Tags</Label>
                        {selectedItem.tags && selectedItem.tags.length > 0 ? (
                            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '5px' }}>
                                {selectedItem.tags.map((tag, idx) => (
                                    <span key={idx} style={{ background: '#f3f2f1', padding: '2px 8px', borderRadius: '4px', fontSize: '12px' }}>
                                        {tag}
                                    </span>
                                ))}
                            </div>
                        ) : (
                            <Text style={{ fontStyle: 'italic', color: '#666' }}>No tags available</Text>
                        )}
                    </div>

                    <Separator />

                    <Text variant="small" styles={{ root: { textAlign: 'center', marginTop: 'auto', color: '#999' } }}>
                        Item {selectedIndex + 1} of {items.length}
                    </Text>

                </Stack>
              </Stack>
            </div>
          )}
        </Modal>

      </Stack>
    );
  }
}