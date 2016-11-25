<?php

/**
 * This is the model class for table "masters".
 *
 * The followings are the available columns in table 'masters':
 * @property string $MasterID
 * @property string $Description
 * @property string $ShortDescription
 * @property integer $Enabled
 *
 * The followings are the available model relations:
 * @property PersonsMasters[] $personsMasters
 */
class Masters extends CActiveRecord
{
	/**
	 * @return string the associated database table name
	 */
	public function tableName()
	{
		return 'masters';
	}

	/**
	 * @return array validation rules for model attributes.
	 */
	public function rules()
	{
		// NOTE: you should only define rules for those attributes that
		// will receive user inputs.
		return array(
			array('Description, Enabled', 'required'),
			array('Enabled', 'numerical', 'integerOnly'=>true),
			array('Description, ShortDescription', 'length', 'max'=>255),
			// The following rule is used by search().
			// @todo Please remove those attributes that should not be searched.
			array('MasterID, Description, ShortDescription, Enabled', 'safe', 'on'=>'search'),
		);
	}

	/**
	 * @return array relational rules.
	 */
	public function relations()
	{
		// NOTE: you may need to adjust the relation name and the related
		// class name for the relations automatically generated below.
		return array(
			'personsMasters' => array(self::HAS_MANY, 'PersonsMasters', 'MasterID'),
		);
	}

	/**
	 * @return array customized attribute labels (name=>label)
	 */
	public function attributeLabels()
	{
		return array(
			'MasterID' => 'Master',
			'Description' => 'Description',
			'ShortDescription' => 'Short Description',
			'Enabled' => 'Enabled',
		);
	}

	/**
	 * Retrieves a list of models based on the current search/filter conditions.
	 *
	 * Typical usecase:
	 * - Initialize the model fields with values from filter form.
	 * - Execute this method to get CActiveDataProvider instance which will filter
	 * models according to data in model fields.
	 * - Pass data provider to CGridView, CListView or any similar widget.
	 *
	 * @return CActiveDataProvider the data provider that can return the models
	 * based on the search/filter conditions.
	 */
	public function search()
	{
		// @todo Please modify the following code to remove attributes that should not be searched.

		$criteria=new CDbCriteria;

		$criteria->compare('MasterID',$this->MasterID,true);
		$criteria->compare('Description',$this->Description,true);
		$criteria->compare('ShortDescription',$this->ShortDescription,true);
		$criteria->compare('Enabled',$this->Enabled);

		return new CActiveDataProvider($this, array(
			'criteria'=>$criteria,
		));
	}

	/**
	 * Returns the static model of the specified AR class.
	 * Please note that you should have this exact method in all your CActiveRecord descendants!
	 * @param string $className active record class name.
	 * @return Masters the static model class
	 */
	public static function model($className=__CLASS__)
	{
		return parent::model($className);
	}
}
